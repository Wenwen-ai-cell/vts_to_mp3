#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
vts2mp3 —— 把「.vts」文件批量转换成 MP3。

这类 .vts 实际上并不是 DVD 的 VTS，而是伪装成 .vts 的 MP4/M4A 容器：
里面通常只有一条 AAC 音轨，外加一张封面图。有些文件的开头部分被覆盖过，
导致 ftyp/mdat 结构头损坏，用播放器或普通转换工具都打不开。

本工具的做法是绕开损坏的部分：
    解析残留的 moov -> 取出逐个 AAC 帧 -> 剔除损坏帧 -> 重新封装成标准 M4A
    -> 交给 Windows 自带的 Media Foundation 解码并编码成 MP3 -> 写入 ID3 标签

特点
  * 不需要 ffmpeg，不需要联网：转码走系统内置的 Media Foundation（WinRT）
  * 能救回开头被覆盖的文件（自动跳过损坏帧）
  * 如果同目录/指定目录里能找到原始录屏，可以用它补回开头损坏的那几秒
  * 自动把标题、专辑、封面写进 MP3
  * 自带网页界面，也可以命令行批量跑

用法
    python vts2mp3.py                          启动网页界面
    python vts2mp3.py --cli <目录或文件...>      命令行批量转换
    python vts2mp3.py --cli <目录> -o <输出目录> --bitrate 320 --src <源录像目录>
"""

import argparse
import json
import os
import re
import struct
import subprocess
import sys
import tempfile
import threading
import time
import traceback
import webbrowser
from collections import Counter
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, parse_qs

APP_NAME = "vts2mp3"
VERSION = "1.0"

# ===========================================================================
# 内嵌的转码引擎（PowerShell 5.1 + WinRT MediaTranscoder）
#
# PowerShell 5.1 无法 await IAsyncAction，所以这里先 await 掉
# PrepareFileTranscodeAsync，再触发 TranscodeAsync 并轮询输出文件大小，
# 直到它不再增长为止。
# ===========================================================================
ENGINE_PS1 = r'''
param(
  [Parameter(Mandatory=$true)][string]$InPath,
  [Parameter(Mandatory=$true)][string]$OutPath,
  [ValidateSet('mp3','wav')][string]$Format = 'mp3',
  [double]$ExpectedSeconds = 0,
  [int]$Bitrate = 192000,
  [int]$StableSeconds = 10,
  [int]$MaxSeconds = 1800
)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
function Say($s) { [Console]::Out.WriteLine($s); [Console]::Out.Flush() }
try {
    Add-Type -AssemblyName System.Runtime.WindowsRuntime | Out-Null
    $asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
        $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and
        $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
    function Await($op, $type) {
        $m = $asTaskGeneric.MakeGenericMethod($type)
        $t = $m.Invoke($null, @($op)); $t.Wait(-1) | Out-Null; $t.Result
    }
    [Windows.Storage.StorageFile,   Windows.Storage, ContentType=WindowsRuntime] | Out-Null
    [Windows.Storage.StorageFolder, Windows.Storage, ContentType=WindowsRuntime] | Out-Null
    [Windows.Media.Transcoding.MediaTranscoder, Windows.Media, ContentType=WindowsRuntime] | Out-Null
    [Windows.Media.MediaProperties.MediaEncodingProfile, Windows.Media, ContentType=WindowsRuntime] | Out-Null
    [Windows.Media.MediaProperties.AudioEncodingQuality, Windows.Media, ContentType=WindowsRuntime] | Out-Null

    if (Test-Path $OutPath) { Remove-Item $OutPath -Force }
    $src    = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($InPath)) ([Windows.Storage.StorageFile])
    $outDir = [System.IO.Path]::GetDirectoryName($OutPath)
    $outNm  = [System.IO.Path]::GetFileName($OutPath)
    if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
    $folder = Await ([Windows.Storage.StorageFolder]::GetFolderFromPathAsync($outDir)) ([Windows.Storage.StorageFolder])
    $dst    = Await ($folder.CreateFileAsync($outNm, [Windows.Storage.CreationCollisionOption]::ReplaceExisting)) ([Windows.Storage.StorageFile])

    $profClass = [Windows.Media.MediaProperties.MediaEncodingProfile]
    if ($Format -eq 'wav') { $profile = $profClass::CreateWav([Windows.Media.MediaProperties.AudioEncodingQuality]::High) }
    else                   { $profile = $profClass::CreateMp3([Windows.Media.MediaProperties.AudioEncodingQuality]::High) }

    $transcoder = New-Object Windows.Media.Transcoding.MediaTranscoder
    $prep = Await ($transcoder.PrepareFileTranscodeAsync($src, $dst, $profile)) ([Windows.Media.Transcoding.PrepareTranscodeResult])
    if (-not $prep.CanTranscode) { Say "STATUS: PREPARE_FAILED"; Say "REASON: $($prep.FailureReason)"; exit 2 }

    $expected = 0
    if ($Format -eq 'mp3' -and $ExpectedSeconds -gt 0) { $expected = [int]($ExpectedSeconds * $Bitrate / 8.0 * 0.97) }
    $null = $prep.TranscodeAsync()

    $sw = [System.Diagnostics.Stopwatch]::StartNew(); $prev = -1; $stable = 0
    while ($sw.Elapsed.TotalSeconds -lt $MaxSeconds) {
        Start-Sleep -Milliseconds 500
        $len = 0; if (Test-Path $OutPath) { $len = (Get-Item $OutPath).Length }
        if ($len -eq $prev -and $len -gt 0) { $stable++ } else { $stable = 0 }
        $prev = $len
        if ($stable -ge ($StableSeconds * 2)) { break }
    }
    if ($prev -le 0) { Say "STATUS: EMPTY"; exit 4 }
    Say "STATUS: OK"; Say ("SIZE: {0}" -f $prev); Say ("ELAPSED: {0:N1}" -f $sw.Elapsed.TotalSeconds); exit 0
}
catch { Say "STATUS: ERROR"; Say "REASON: $($_.Exception.Message)"; exit 3 }
'''

# ===========================================================================
# MP4 / AAC 解析与重新封装
# ===========================================================================
AAC_RATES = [96000, 88200, 64000, 48000, 44100, 32000, 24000, 22050,
             16000, 12000, 11025, 8000, 7350]


def iter_atoms(buf, start, end):
    off = start
    while off + 8 <= end:
        size = struct.unpack(">I", buf[off:off + 4])[0]
        typ = buf[off + 4:off + 8]
        hdr = 8
        if size == 1:
            if off + 16 > end:
                return
            size = struct.unpack(">Q", buf[off + 8:off + 16])[0]
            hdr = 16
        elif size == 0:
            size = end - off
        if size < hdr or off + size > end:
            return
        yield off, size, typ, hdr
        off += size


def find_child(buf, start, end, typ):
    for off, size, t, hdr in iter_atoms(buf, start, end):
        if t == typ:
            return off, size, hdr
    return None


def atom(typ, payload):
    return struct.pack(">I", 8 + len(payload)) + typ + payload


def fullbox(typ, payload, version=0, flags=0):
    return atom(typ, struct.pack(">B", version) + struct.pack(">I", flags)[1:] + payload)


def find_moov(data):
    """定位结构有效的 moov（这类文件把它放在文件末尾）。"""
    n = len(data)
    idx = data.rfind(b"moov")
    while idx != -1:
        start = idx - 4
        if start >= 0:
            size = struct.unpack(">I", data[start:start + 4])[0]
            if 8 <= size <= n - start:
                kids = {t for _, _, t, _ in iter_atoms(data, start + 8, start + size)}
                if b"trak" in kids or b"mvhd" in kids:
                    return start, size
        idx = data.rfind(b"moov", 0, idx)
    return None


def _parse_stbl(data, off, end):
    out = {}
    for a, size, typ, hdr in iter_atoms(data, off, end):
        p = a + hdr
        if typ == b"stsd":
            out["stsd"] = data[p:a + size]
        elif typ == b"stts":
            cnt = struct.unpack(">I", data[p + 4:p + 8])[0]
            out["stts"] = [struct.unpack(">II", data[p + 8 + 8 * i:p + 16 + 8 * i]) for i in range(cnt)]
        elif typ == b"stsc":
            cnt = struct.unpack(">I", data[p + 4:p + 8])[0]
            out["stsc"] = [struct.unpack(">III", data[p + 8 + 12 * i:p + 20 + 12 * i]) for i in range(cnt)]
        elif typ == b"stsz":
            uniform, cnt = struct.unpack(">II", data[p + 4:p + 12])
            out["sizes"] = ([uniform] * cnt if uniform else
                            list(struct.unpack(">%dI" % cnt, data[p + 12:p + 12 + 4 * cnt])))
        elif typ in (b"stco", b"co64"):
            cnt = struct.unpack(">I", data[p + 4:p + 8])[0]
            fmt = ">%dI" % cnt if typ == b"stco" else ">%dQ" % cnt
            step = 4 if typ == b"stco" else 8
            out["chunks"] = list(struct.unpack(fmt, data[p + 8:p + 8 + step * cnt]))
    return out


def _read_desc_len(buf, i, end):
    ln, j = 0, i
    for _ in range(4):
        if j >= end:
            return None, None
        ln = (ln << 7) | (buf[j] & 0x7F)
        cont = buf[j] & 0x80
        j += 1
        if not cont:
            return ln, j
    return None, None


def find_asc(stsd):
    """从 stsd 里取出 AudioSpecificConfig。"""
    e = stsd.find(b"esds")
    if e == -1:
        return None
    end = min(len(stsd), e + 300)
    i = e + 8
    while i < end - 1:
        tag = stsd[i]
        ln, j = _read_desc_len(stsd, i + 1, end)
        if ln is None:
            break
        if tag == 0x05:
            return stsd[j:j + ln]
        if tag == 0x03:
            i = j + 3
        elif tag == 0x04:
            i = j + 13
        else:
            i = j + ln
    for pat in (b"\x05\x80\x80\x80", b"\x05\x80\x80", b"\x05"):
        k = stsd.find(pat, e, end)
        while k != -1:
            ln, j = _read_desc_len(stsd, k + 1, end)
            if ln and 2 <= ln <= 8 and j + ln <= len(stsd):
                cand = stsd[j:j + ln]
                if (cand[0] >> 3) in (1, 2, 5, 29):
                    return cand
            k = stsd.find(pat, k + 1, end)
    return None


def parse_audio_track(data, moov_off, moov_size):
    """找到第一条音频轨（smhd），返回它的参数和采样表。"""
    for off, size, typ, hdr in iter_atoms(data, moov_off + 8, moov_off + moov_size):
        if typ != b"trak":
            continue
        mdia = find_child(data, off + hdr, off + size, b"mdia")
        if not mdia:
            continue
        minf = find_child(data, mdia[0] + mdia[2], mdia[0] + mdia[1], b"minf")
        if not minf:
            continue
        if not find_child(data, minf[0] + minf[2], minf[0] + minf[1], b"smhd"):
            continue
        mdhd = find_child(data, mdia[0] + mdia[2], mdia[0] + mdia[1], b"mdhd")
        timescale = 48000
        if mdhd:
            p = mdhd[0] + mdhd[2]
            ver = data[p]
            timescale = struct.unpack(">I", data[p + (20 if ver == 1 else 12):
                                                 p + (24 if ver == 1 else 16)])[0]
        stbl = find_child(data, minf[0] + minf[2], minf[0] + minf[1], b"stbl")
        if not stbl:
            continue
        tables = _parse_stbl(data, stbl[0] + stbl[2], stbl[0] + stbl[1])
        if "sizes" not in tables or "chunks" not in tables:
            continue

        asc, channels, rate = None, 2, timescale or 48000
        stsd = tables.get("stsd")
        if stsd:
            asc = find_asc(stsd)
            m = stsd.find(b"mp4a")
            if m != -1 and m + 32 <= len(stsd):
                channels = struct.unpack(">H", stsd[m + 20:m + 22])[0] or channels
                r = struct.unpack(">I", stsd[m + 28:m + 32])[0] >> 16
                if r:
                    rate = r
        if asc and len(asc) >= 2:
            sfi = ((asc[0] & 0x07) << 1) | (asc[1] >> 7)
            ch = (asc[1] >> 3) & 0x0F
            if ch:
                channels = ch
            if sfi < len(AAC_RATES):
                rate = AAC_RATES[sfi]
        return {"timescale": timescale, "rate": rate, "channels": channels,
                "asc": asc, "tables": tables}
    return None


def extract_frames(data, track):
    """按解码顺序把所有采样（AAC 帧）拼起来。"""
    t = track["tables"]
    sizes, chunks = t["sizes"], t["chunks"]
    stsc = t.get("stsc") or [(1, len(sizes), 1)]
    spc = {}
    for i, (first, per, _d) in enumerate(stsc):
        last = stsc[i + 1][0] - 1 if i + 1 < len(stsc) else len(chunks)
        for c in range(first, last + 1):
            spc[c] = per
    frames = []
    for c in range(1, len(chunks) + 1):
        off = chunks[c - 1]
        for _ in range(spc.get(c, 0)):
            if len(frames) >= len(sizes):
                break
            s = sizes[len(frames)]
            if off + s > len(data):
                return frames
            frames.append(data[off:off + s])
            off += s
    return frames


def dominant_first_byte(frames):
    """这类编码器产出的合法 AAC 帧，首字节几乎总是同一个值（通常是 0x01）。"""
    c = Counter(f[0] for f in frames if f)
    if not c:
        return None, 0.0
    byte, cnt = c.most_common(1)[0]
    return byte, cnt / sum(c.values())


def usable_segment(frames, valid_byte, min_frac=0.6, gap=3):
    """挑出最长的一段完好数据。

    首字节等于主导字节的帧视为完好；连续 < gap 个异常帧当作噪声容忍，
    连续 >= gap 个异常帧则判定为「被覆盖的区域」，取最长的无损伤区间。
    """
    if not frames:
        return 0, 0, False, []
    ok = [bool(f) and f[0] == valid_byte for f in frames]
    trusted = (sum(ok) / len(ok)) >= min_frac
    zones, i = [], 0
    while i < len(ok):
        if not ok[i]:
            j = i
            while j < len(ok) and not ok[j]:
                j += 1
            if j - i >= gap:
                zones.append((i, j))
            i = j
        else:
            i += 1
    if not zones:
        return 0, len(frames), trusted, []
    segs, prev = [], 0
    for a, b in zones:
        if a > prev:
            segs.append((prev, a))
        prev = b
    if prev < len(frames):
        segs.append((prev, len(frames)))
    if not segs:
        return 0, 0, trusted, zones
    best = max(segs, key=lambda s: s[1] - s[0])
    return best[0], best[1], trusted, zones


def _desc_len(n):
    if n < 0x80:
        return bytes([n])
    if n < 0x4000:
        return bytes([0x80 | (n >> 7), n & 0x7F])
    return bytes([0x80 | ((n >> 14) & 0x7F), 0x80 | ((n >> 7) & 0x7F), n & 0x7F])


def _mp4a_entry(asc, channels, rate):
    esds_body = (b"\x03" + _desc_len(0x22) + b"\x00\x01\x00" +
                 b"\x04" + _desc_len(0x14) + b"\x40\x15\x00\x00\x00" +
                 struct.pack(">II", 0x000260EC, 0x000260E8) +
                 b"\x05" + _desc_len(len(asc)) + asc +
                 b"\x06" + _desc_len(1) + b"\x02")
    p = (b"\x00" * 6 + struct.pack(">H", 1) +
         struct.pack(">HH", 0, 0) + b"\x00" * 4 +
         struct.pack(">H", channels) + struct.pack(">H", 16) +
         struct.pack(">HH", 0, 0) + struct.pack(">I", rate << 16))
    return atom(b"mp4a", p + fullbox(b"esds", esds_body))


def build_m4a(frames, asc, channels=2, rate=48000):
    """把 AAC 帧封装成一个最小但完全合法的 M4A。"""
    if not asc:
        # 兜底：AAC-LC，取样率与声道沿用轨道里的值
        sfi = AAC_RATES.index(rate) if rate in AAC_RATES else 3
        asc = bytes([(2 << 3) | (sfi >> 1), ((sfi & 1) << 7) | (channels << 3)])
    total = len(frames) * 1024
    dur_ms = int(round(total * 1000.0 / rate))
    matrix = struct.pack(">9i", 0x10000, 0, 0, 0, 0x10000, 0, 0, 0, 0x40000000)

    mvhd = fullbox(b"mvhd", struct.pack(">IIII", 0, 0, 1000, dur_ms) +
                   struct.pack(">I", 0x00010000) + struct.pack(">H", 0x0100) +
                   b"\x00" * 10 + matrix + b"\x00" * 24 + struct.pack(">I", 2))
    tkhd = fullbox(b"tkhd", struct.pack(">IIIII", 0, 0, 1, 0, dur_ms) +
                   b"\x00" * 8 + struct.pack(">HHH", 0, 0, 0x0100) + b"\x00\x00" +
                   matrix + struct.pack(">II", 0, 0), flags=7)
    mdhd = fullbox(b"mdhd", struct.pack(">IIII", 0, 0, rate, total) + struct.pack(">HH", 0x55C4, 0))
    hdlr = fullbox(b"hdlr", b"\x00" * 4 + b"soun" + b"\x00" * 12 + b"SoundHandler\x00")
    smhd = fullbox(b"smhd", struct.pack(">HH", 0, 0))
    dref = fullbox(b"dref", struct.pack(">I", 1) + atom(b"url ", b"\x00\x00\x00\x01"))
    dinf = atom(b"dinf", dref)
    stsd = fullbox(b"stsd", struct.pack(">I", 1) + _mp4a_entry(asc, channels, rate))
    stts = fullbox(b"stts", struct.pack(">I", 1) + struct.pack(">II", len(frames), 1024))
    stsc = fullbox(b"stsc", struct.pack(">I", 1) + struct.pack(">III", 1, len(frames), 1))
    stsz = fullbox(b"stsz", struct.pack(">II", 0, len(frames)) +
                   b"".join(struct.pack(">I", len(f)) for f in frames))

    ftyp = atom(b"ftyp", b"M4A " + struct.pack(">I", 0) + b"M4A isommp42")
    # moov 大小恒定，所以可以精确算出 mdat 的实际数据偏移
    probe = atom(b"moov", mvhd + atom(b"trak", tkhd + atom(b"mdia", mdhd + hdlr +
                atom(b"minf", smhd + dinf + atom(b"stbl", stsd + stts + stsc + stsz +
                     fullbox(b"stco", struct.pack(">II", 1, 0)))))))
    data_off = len(ftyp) + len(probe) + 8
    stco = fullbox(b"stco", struct.pack(">II", 1, data_off))
    stbl = atom(b"stbl", stsd + stts + stsc + stsz + stco)
    moov = atom(b"moov", mvhd + atom(b"trak", tkhd + atom(b"mdia", mdhd + hdlr +
                atom(b"minf", smhd + dinf + stbl))))
    return ftyp + moov + atom(b"mdat", b"".join(frames)), data_off, total / rate


def read_ilst(data, moov_off, moov_size):
    """读取 moov > udta > meta > ilst 里的标题/专辑/封面。"""
    meta = {"title": None, "album": None, "artist": None, "cover": None}
    udta = find_child(data, moov_off + 8, moov_off + moov_size, b"udta")
    if not udta:
        return meta
    m = find_child(data, udta[0] + udta[2], udta[0] + udta[1], b"meta")
    if not m:
        return meta
    ilst = find_child(data, m[0] + m[2] + 4, m[0] + m[1], b"ilst")
    if not ilst:
        return meta
    keys = {b"\xa9nam": "title", b"\xa9alb": "album", b"\xa9ART": "artist"}
    for off, size, typ, hdr in iter_atoms(data, ilst[0] + ilst[2], ilst[0] + ilst[1]):
        if typ not in keys and typ != b"covr":
            continue
        d = find_child(data, off + hdr, off + size, b"data")
        if not d:
            continue
        body = data[d[0] + d[2] + 8:d[0] + d[1]]
        if typ == b"covr":
            meta["cover"] = body
        else:
            try:
                meta[keys[typ]] = body.decode("utf-8")
            except UnicodeDecodeError:
                pass
    return meta


def build_id3(meta):
    """ID3v2.3 标签：标题、专辑、艺术家，以及内嵌封面。"""
    def syncsafe(n):
        return bytes([(n >> 21) & 0x7F, (n >> 14) & 0x7F, (n >> 7) & 0x7F, n & 0x7F])

    def text(fid, s):
        body = b"\x01" + "\ufeff".encode("utf-16-be") + s.encode("utf-16-be")
        return fid + struct.pack(">I", len(body)) + b"\x00\x00" + body

    out = b""
    for fid, key in ((b"TIT2", "title"), (b"TALB", "album"), (b"TPE1", "artist")):
        if meta.get(key):
            out += text(fid, meta[key])
    out += text(b"TSSE", APP_NAME)
    cover = meta.get("cover")
    if cover:
        mime = b"image/png" if cover[:4] == b"\x89PNG" else b"image/jpeg"
        body = b"\x00" + mime + b"\x00" + b"\x03" + b"\x00" + cover
        out += b"APIC" + struct.pack(">I", len(body)) + b"\x00\x00" + body
    return (b"ID3\x03\x00\x00" + syncsafe(len(out)) + out) if out else b""


# ===========================================================================
# 转换流程
# ===========================================================================
def analyze_file(path):
    """读入并分析一个 .vts，返回 (摘要, 原始数据)。"""
    info = {"path": path, "name": os.path.basename(path), "ok": False,
            "size": os.path.getsize(path), "error": None}
    try:
        with open(path, "rb") as fh:
            data = fh.read()
    except OSError as e:
        info["error"] = "无法读取：%s" % e
        return info, None

    mo = find_moov(data)
    if not mo:
        info["error"] = "找不到 moov，不是同类 .vts（可能是真正的 DVD VTS 或其他格式）"
        return info, None
    track = parse_audio_track(data, mo[0], mo[1])
    if not track:
        info["error"] = "文件里没有音频轨道"
        return info, None

    frames = extract_frames(data, track)
    if not frames:
        info["error"] = "采样表为空，取不到音频数据"
        return info, None

    vbyte, frac = dominant_first_byte(frames)
    start, end, trusted, zones = usable_segment(frames, vbyte)
    rate = track["rate"] or 48000
    meta = read_ilst(data, mo[0], mo[1])

    info.update({
        "ok": True, "rate": rate, "channels": track["channels"],
        "frames": len(frames), "frames_used": end - start,
        "valid_byte_is_trusted": trusted,
        "dropped_frames": start,
        "dropped_seconds": start * 1024.0 / rate,
        "total_seconds": len(frames) * 1024.0 / rate,
        "duration": (end - start) * 1024.0 / rate,
        "damaged": start > 0 or (len(frames) - end) > 0,
        "title": meta.get("title"), "album": meta.get("album"),
        "has_cover": bool(meta.get("cover")),
        "seg_start": start, "seg_end": end,
    })
    if info["frames_used"] * 1024.0 / rate < 1.0:
        info["ok"] = False
        info["error"] = "可用的音频不足 1 秒，文件可能损坏过于严重"
    return info, {"data": data, "track": track, "frames": frames, "meta": meta}


SRC_EXTS = (".mp4", ".m4a", ".mov", ".mkv", ".mp3", ".vts", ".avi",
            ".flv", ".wmv", ".ts", ".webm", ".3gp", ".m4v")


def find_source_file(vts_path, frames, seg_start, seg_end, search_dirs):
    """找原始录屏，用来补回开头被覆盖的部分。

    只有源文件在完好区间内的帧与 .vts 完全一致时才会采用。
    返回 (路径, 帧列表, 被否决的候选列表)。
    """
    base = os.path.splitext(os.path.basename(vts_path))[0]
    # .vts 名字通常是「源录像名 + 一串随机数字」，所以用前缀关系来配对，
    # 真正是否采用由后面的逐帧比对决定。
    def name_matches(stem):
        if stem == base:
            return True
        if len(stem) >= 8 and base.startswith(stem):
            return True
        if len(base) >= 8 and stem.startswith(base):
            return True
        return False

    dirs = [d for d in ([os.path.dirname(vts_path)] + list(search_dirs))
            if d and os.path.isdir(d)]
    cands = []
    for d in dirs:
        try:
            names = os.listdir(d)
        except OSError:
            continue
        for name in names:
            stem, ext = os.path.splitext(name)
            if ext.lower() not in SRC_EXTS or not name_matches(stem):
                continue
            full = os.path.join(d, name)
            if os.path.abspath(full) == os.path.abspath(vts_path):
                continue
            try:
                sz = os.path.getsize(full)
            except OSError:
                continue
            if sz > 800 * 1024 * 1024:
                continue
            cands.append((sz, full))
    cands.sort(reverse=True)                     # 大的更可能是完整录像

    checked = []
    for _sz, full in cands:
        try:
            with open(full, "rb") as fh:
                data = fh.read()
            mo = find_moov(data)
            if not mo:
                continue
            tr = parse_audio_track(data, mo[0], mo[1])
            if not tr:
                continue
            sframes = extract_frames(data, tr)
        except Exception:
            continue
        if len(sframes) == len(frames) and sframes == frames:
            checked.append((full, "与 .vts 内容完全相同"))
            continue
        if len(sframes) < seg_end:
            checked.append((full, "音频长度不足"))
            continue
        if sframes[seg_start:seg_end] != frames[seg_start:seg_end]:
            checked.append((full, "音频内容不一致"))
            continue
        return full, sframes, checked
    return None, None, checked


def run_engine(engine_path, in_path, out_path, fmt="mp3", expected=0.0,
               bitrate=192000, stable=10, timeout=2400):
    """调用 PowerShell 转码引擎，返回 (是否成功, 失败原因, 输出大小)。"""
    ps = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"),
                      "System32", "WindowsPowerShell", "v1.0", "powershell.exe")
    if not os.path.exists(ps):
        ps = "powershell.exe"
    cmd = [ps, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
           "-File", engine_path, "-InPath", in_path, "-OutPath", out_path,
           "-Format", fmt, "-ExpectedSeconds", str(expected),
           "-Bitrate", str(int(bitrate)), "-StableSeconds", str(stable)]
    try:
        r = subprocess.run(cmd, capture_output=True, timeout=timeout)
    except subprocess.TimeoutExpired:
        return False, "转码超时", 0
    except OSError as e:
        return False, "无法启动 PowerShell：%s" % e, 0
    out = (r.stdout or b"").decode("utf-8", "replace")
    err = (r.stderr or b"").decode("utf-8", "replace")
    size = os.path.getsize(out_path) if os.path.exists(out_path) else 0
    if "STATUS: OK" in out and size > 0:
        return True, "", size
    reason = ""
    for line in out.splitlines():
        if line.startswith("REASON:"):
            reason = line[7:].strip()
    if not reason:
        tail = [l for l in out.strip().splitlines() if l.strip()]
        reason = (tail[-1] if tail else "") or err.strip()[:200] or "转码失败"
    return False, reason, size


def make_tmpdir(explicit=None):
    """挑一个确实可写的临时目录。

    刻意不用 tempfile.mkdtemp：它建出来的目录权限在部分受限环境下无法写入，
    这里改用 os.makedirs 并加唯一后缀。
    """
    bases = [explicit] if explicit else []
    bases += [tempfile.gettempdir(), os.getcwd()]
    for base in bases:
        if not base:
            continue
        d = os.path.join(base, "vts2mp3_%d_%d" % (os.getpid(), int(time.time() * 1000) % 100000))
        try:
            os.makedirs(d, exist_ok=True)
            probe = os.path.join(d, "probe.tmp")
            with open(probe, "w") as fh:
                fh.write("x")
            os.remove(probe)
            return d
        except OSError:
            try:
                os.rmdir(d)
            except OSError:
                pass
            continue
    raise RuntimeError("找不到可写的临时目录，请用 --tmpdir 指定一个")


class Converter:
    """单文件转换器，启动时把内嵌的引擎脚本落盘。"""

    def __init__(self, bitrate=256, tmpdir=None):
        self.bitrate = int(bitrate) * 1000
        self.tmpdir = make_tmpdir(tmpdir)
        self.engine = os.path.join(self.tmpdir, "engine.ps1")
        with open(self.engine, "w", encoding="utf-8-sig") as fh:
            fh.write(ENGINE_PS1)

    def convert(self, src_path, out_path, search_dirs=(), log=print):
        """转换单个文件，返回结果字典。"""
        res = {"input": src_path, "output": out_path, "ok": False, "error": None,
               "used_source": None, "duration": 0.0, "skipped_seconds": 0.0,
               "output_size": 0}
        info, blob = analyze_file(src_path)
        if not info["ok"]:
            res["error"] = info["error"]
            return res

        track, frames = blob["track"], blob["frames"]
        meta = blob["meta"]
        a, b = info["seg_start"], info["seg_end"]
        seg_frames = frames[a:b]

        src_file, src_frames, checked = find_source_file(src_path, frames, a, b, search_dirs)
        if src_file and src_frames:
            seg_frames = src_frames[:b]
            res["used_source"] = src_file
            log("   用源录像补回开头：%s" % os.path.basename(src_file))
        else:
            for f, why in checked:
                log("   跳过候选源 %s（%s）" % (os.path.basename(f), why))
            if a > 0:
                res["skipped_seconds"] = info["dropped_seconds"]
                log("   注意：开头 %.1f 秒已被覆盖，将从第 %.1f 秒处开始转"
                    % (info["dropped_seconds"], info["dropped_seconds"]))

        if not seg_frames:
            res["error"] = "没有可用的音频帧"
            return res

        m4a_bytes, _off, duration = build_m4a(seg_frames, track["asc"],
                                              track["channels"], track["rate"] or 48000)
        res["duration"] = duration

        tag = abs(hash(src_path)) % 10 ** 10
        tmp_m4a = os.path.join(self.tmpdir, "in_%d.m4a" % tag)
        tmp_mp3 = os.path.join(self.tmpdir, "out_%d.mp3" % tag)
        try:
            with open(tmp_m4a, "wb") as fh:
                fh.write(m4a_bytes)
            ok, why, _size = run_engine(self.engine, tmp_m4a, tmp_mp3, "mp3",
                                        expected=duration, bitrate=self.bitrate)
            if not ok:
                res["error"] = "转码失败：%s" % why
                return res
            with open(tmp_mp3, "rb") as fh:
                audio = fh.read()
        finally:
            for p in (tmp_m4a, tmp_mp3):
                try:
                    os.remove(p)
                except OSError:
                    pass

        id3 = build_id3(meta)
        outdir = os.path.dirname(os.path.abspath(out_path))
        os.makedirs(outdir, exist_ok=True)
        with open(out_path, "wb") as fh:
            fh.write(id3)
            fh.write(audio)
        res["ok"] = True
        res["output_size"] = len(id3) + len(audio)
        return res


# ===========================================================================
# 命令行
# ===========================================================================
def collect_inputs(paths):
    out = []
    for p in paths:
        if os.path.isdir(p):
            for name in sorted(os.listdir(p)):
                if name.lower().endswith(".vts"):
                    out.append(os.path.join(p, name))
        elif os.path.isfile(p):
            out.append(p)
    return out


def cli(args):
    files = collect_inputs(args.inputs)
    if not files:
        print("没有找到任何 .vts 文件")
        return 1
    outdir = args.output or os.path.join(os.path.dirname(os.path.abspath(files[0])), "mp3")
    conv = Converter(bitrate=args.bitrate, tmpdir=args.tmpdir)
    search = [d for d in (args.src or []) if d]

    print("%s v%s —— 共 %d 个文件 -> %s\n" % (APP_NAME, VERSION, len(files), outdir))
    ok = 0
    for i, f in enumerate(files, 1):
        print("[%d/%d] %s" % (i, len(files), os.path.basename(f)))
        out = os.path.join(outdir, os.path.splitext(os.path.basename(f))[0] + ".mp3")
        try:
            r = conv.convert(f, out, search, log=lambda s: print(s))
        except Exception as e:
            r = {"ok": False, "error": "%s: %s" % (type(e).__name__, e)}
            if args.verbose:
                traceback.print_exc()
        if r["ok"]:
            ok += 1
            print("   -> 完成 %.1f 秒, %.0f KB\n" % (r["duration"], r["output_size"] / 1024))
        else:
            print("   -> 失败：%s\n" % r["error"])
    print("完成：%d/%d 成功" % (ok, len(files)))
    return 0 if ok == len(files) else 2


# ===========================================================================
# 网页界面
# ===========================================================================
PAGE = r"""<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8">
<title>vts2mp3 - .vts 转 MP3</title>
<style>
 :root{--bg:#0f1216;--panel:#171c23;--line:#252c36;--fg:#e6ebf2;--dim:#8b98a9;--acc:#4c8dff;--ok:#3fb950;--warn:#d29922;--err:#f85149}
 *{box-sizing:border-box}
 body{margin:0;background:var(--bg);color:var(--fg);font:14px/1.55 "Segoe UI",system-ui,"Microsoft YaHei",sans-serif}
 .wrap{max-width:1080px;margin:0 auto;padding:28px 20px 60px}
 h1{font-size:22px;margin:0 0 4px} .sub{color:var(--dim);margin-bottom:22px;font-size:13px}
 .card{background:var(--panel);border:1px solid var(--line);border-radius:10px;padding:16px 18px;margin-bottom:16px}
 label{display:block;font-size:12px;color:var(--dim);margin-bottom:5px}
 .row{display:flex;gap:10px;align-items:flex-end;flex-wrap:wrap}
 .grow{flex:1;min-width:260px}
 input,select{width:100%;background:#0d1117;border:1px solid var(--line);color:var(--fg);
   border-radius:7px;padding:9px 11px;font:inherit;font-size:13px}
 input:focus,select:focus{outline:none;border-color:var(--acc)}
 button{background:var(--acc);color:#fff;border:0;border-radius:7px;padding:10px 18px;
   font:inherit;font-weight:600;cursor:pointer}
 button.ghost{background:#222a35} button:disabled{opacity:.45;cursor:default}
 table{width:100%;border-collapse:collapse;font-size:13px}
 th,td{text-align:left;padding:8px 10px;border-bottom:1px solid var(--line);vertical-align:top}
 th{color:var(--dim);font-weight:500;font-size:12px}
 .tag{display:inline-block;padding:1px 7px;border-radius:5px;font-size:11px;white-space:nowrap}
 .t-ok{background:rgba(63,185,80,.15);color:var(--ok)} .t-warn{background:rgba(210,153,34,.15);color:var(--warn)}
 .t-err{background:rgba(248,81,73,.15);color:var(--err)}
 #log{background:#0d1117;border:1px solid var(--line);border-radius:8px;padding:12px;
   height:210px;overflow:auto;font:12px/1.6 Consolas,monospace;white-space:pre-wrap;color:#b9c4d2}
 .bar{height:6px;background:#0d1117;border-radius:4px;overflow:hidden;margin-top:12px}
 .bar>i{display:block;height:100%;width:0;background:var(--acc);transition:width .3s}
 .hint{color:var(--dim);font-size:12px;margin-top:8px}
 .stat{display:flex;gap:22px;color:var(--dim);font-size:12px;margin-top:10px;flex-wrap:wrap}
 .stat b{color:var(--fg);font-size:15px;display:block;font-weight:600}
</style></head><body><div class="wrap">
<h1>vts2mp3 <span style="color:var(--dim);font-size:13px;font-weight:400">v__VERSION__</span></h1>
<div class="sub">把 .vts（实为 MP4/AAC 容器）批量转成 MP3 - 走 Windows 自带 Media Foundation，无需 ffmpeg 与联网</div>

<div class="card">
  <div class="row">
    <div class="grow"><label>要转换的文件夹（里面放 .vts 文件）</label>
      <input id="dir" placeholder="例如 D:\录像\vts"></div>
    <button class="ghost" id="btnScan">扫描</button>
  </div>
  <div class="row" style="margin-top:12px">
    <div class="grow"><label>输出到（留空 = 输入文件夹下的 mp3 子目录）</label>
      <input id="out" placeholder="留空自动"></div>
    <div style="width:150px"><label>比特率</label>
      <select id="br"><option>128</option><option>192</option><option selected>256</option><option>320</option></select></div>
    <div style="width:150px"><label>并行数</label>
      <select id="jobs"><option>1</option><option selected>2</option><option>3</option><option>4</option></select></div>
  </div>
  <div class="row" style="margin-top:12px">
    <div class="grow"><label>原始录屏所在文件夹（可选，用于补回开头被覆盖的几秒）</label>
      <input id="src" placeholder="留空 = 只在 .vts 同目录里找"></div>
    <button id="btnRun" disabled>开始转换</button>
  </div>
  <div class="hint">提示：扫描后可以看到每个文件丢失了多少开头、以及能否用源录像补回。</div>
</div>

<div class="card" id="listCard" style="display:none">
  <table><thead><tr><th style="width:42%">文件</th><th>状态</th><th>时长</th><th>封面 / 标题</th></tr></thead>
  <tbody id="rows"></tbody></table>
  <div class="stat" id="stat"></div>
  <div class="bar"><i id="prog"></i></div>
</div>

<div class="card"><label style="margin-bottom:8px">日志</label><div id="log">等待开始…</div></div>
</div>
<script>
const $=s=>document.querySelector(s);
let files=[],job=null,timer=null;
function log(s){const l=$('#log');l.textContent+=s+"\n";l.scrollTop=l.scrollHeight;}
function api(u,b){return fetch(u,{method:b?'POST':'GET',headers:{'Content-Type':'application/json'},
  body:b?JSON.stringify(b):undefined}).then(r=>r.json());}
function fmt(n){return n<1024?n+" B":n<1048576?(n/1024).toFixed(0)+" KB":(n/1048576).toFixed(1)+" MB";}
function esc(s){return String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));}

$('#btnScan').onclick=async()=>{
  const dir=$('#dir').value.trim(); if(!dir){alert('请先填写文件夹路径');return;}
  $('#btnScan').disabled=true; $('#log').textContent='正在扫描…\n';
  try{
    const r=await api('/api/scan',{dir:dir});
    if(r.error){log('错误：'+r.error);$('#btnScan').disabled=false;return;}
    files=r.files; render(); log('扫描到 '+files.length+' 个 .vts 文件');
    $('#btnRun').disabled=files.filter(f=>f.ok).length===0;
  }catch(e){log('请求失败：'+e);}
  $('#btnScan').disabled=false;
};

function render(){
  const tb=$('#rows'); tb.innerHTML='';
  files.forEach(f=>{
    const tr=document.createElement('tr');
    let st;
    if(!f.ok) st='<span class="tag t-err">无法识别</span><div style="color:var(--dim);font-size:12px;margin-top:3px">'+esc(f.error||'')+'</div>';
    else if(f.dropped_seconds>0.5) st='<span class="tag t-warn">开头损坏 '+f.dropped_seconds.toFixed(1)+'s</span><div style="color:var(--dim);font-size:12px;margin-top:3px">可转出 '+f.duration.toFixed(1)+'s</div>';
    else st='<span class="tag t-ok">完好</span>';
    tr.innerHTML='<td>'+esc(f.name)+'<div style="color:var(--dim);font-size:12px">'+fmt(f.size)+'</div></td>'+
      '<td>'+st+'</td><td>'+(f.ok?f.duration.toFixed(1)+' s':'-')+'</td>'+
      '<td style="color:var(--dim)">'+(f.has_cover?'有封面':'无')+(f.title?'<div style="font-size:12px">'+esc(f.title)+'</div>':'')+'</td>';
    tb.appendChild(tr);
  });
  const good=files.filter(f=>f.ok).length;
  $('#listCard').style.display='';
  $('#stat').innerHTML='<div><b>'+files.length+'</b>个文件</div><div><b>'+good+'</b>可转换</div>'+
    '<div><b>'+(files.length-good)+'</b>无法识别</div>';
}

$('#btnRun').onclick=async()=>{
  const payload={dir:$('#dir').value.trim(),out:$('#out').value.trim(),src:$('#src').value.trim(),
                 bitrate:parseInt($('#br').value),jobs:parseInt($('#jobs').value)};
  $('#btnRun').disabled=true; $('#log').textContent='';
  const r=await api('/api/start',payload);
  if(r.error){log('错误：'+r.error);$('#btnRun').disabled=false;return;}
  job=r.id; log('开始转换，共 '+r.total+' 个文件…');
  timer=setInterval(poll,600);
};

async function poll(){
  const s=await api('/api/status?id='+job);
  if(s.error) return;
  $('#prog').style.width=(s.done/s.total*100)+'%';
  if(s.lines&&s.lines.length){ s.lines.forEach(l=>log(l)); s.lines=[]; }
  if(s.finished){ clearInterval(timer); $('#btnRun').disabled=false;
    log('\n全部结束：成功 '+s.ok+' / '+s.total); }
}
</script></body></html>
"""


class Handler(BaseHTTPRequestHandler):
    server_version = "%s/%s" % (APP_NAME, VERSION)

    def log_message(self, *a):
        pass

    def _send(self, obj, code=200):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        u = urlparse(self.path)
        if u.path in ("/", "/index.html"):
            body = PAGE.replace("__VERSION__", VERSION).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        elif u.path == "/api/status":
            jid = (parse_qs(u.query).get("id") or [""])[0]
            job = self.server.jobs.get(jid)
            if not job:
                self._send({"error": "任务不存在"})
            else:
                lines, job["lines"] = job["lines"], []   # 取出即清空
                self._send(dict(job, lines=lines))
        else:
            self._send({"error": "not found"}, 404)

    def do_POST(self):
        n = int(self.headers.get("Content-Length") or 0)
        try:
            req = json.loads(self.rfile.read(n) or b"{}")
        except ValueError:
            req = {}
        u = urlparse(self.path)
        if u.path == "/api/scan":
            self._send(self.server.do_scan(req))
        elif u.path == "/api/start":
            self._send(self.server.do_start(req))
        else:
            self._send({"error": "not found"}, 404)


class App(ThreadingHTTPServer):
    daemon_threads = True
    allow_reuse_address = True

    def __init__(self, addr, tmpdir=None):
        super().__init__(addr, Handler)
        self.jobs = {}
        self.tmpdir = tmpdir

    def do_scan(self, req):
        d = (req.get("dir") or "").strip().strip('"')
        if not os.path.isdir(d):
            return {"error": "文件夹不存在：%s" % d}
        files = []
        for name in sorted(os.listdir(d)):
            if not name.lower().endswith(".vts"):
                continue
            p = os.path.join(d, name)
            if not os.path.isfile(p):
                continue
            try:
                info, _ = analyze_file(p)
            except Exception as e:
                info = {"name": name, "size": os.path.getsize(p), "ok": False,
                        "error": "%s: %s" % (type(e).__name__, e)}
            info.pop("path", None)
            files.append(info)
        return {"files": files}

    def do_start(self, req):
        d = (req.get("dir") or "").strip().strip('"')
        if not os.path.isdir(d):
            return {"error": "文件夹不存在：%s" % d}
        out = (req.get("out") or "").strip().strip('"') or os.path.join(d, "mp3")
        src = (req.get("src") or "").strip().strip('"')
        bitrate = int(req.get("bitrate") or 256)
        jobs = max(1, min(6, int(req.get("jobs") or 2)))

        targets = [os.path.join(d, n) for n in sorted(os.listdir(d))
                   if n.lower().endswith(".vts") and os.path.isfile(os.path.join(d, n))]
        if not targets:
            return {"error": "该文件夹里没有 .vts 文件"}

        jid = "job%d" % int(time.time() * 1000)
        self.jobs[jid] = {"total": len(targets), "done": 0, "ok": 0,
                          "lines": [], "finished": False}
        conv = Converter(bitrate=bitrate, tmpdir=self.tmpdir)
        search = [src] if src else []

        def log(line):
            self.jobs[jid]["lines"].append(line)

        def work(path):
            outfile = os.path.join(out, os.path.splitext(os.path.basename(path))[0] + ".mp3")
            log("[%s]" % os.path.basename(path))
            try:
                r = conv.convert(path, outfile, search, log=log)
                if r["ok"]:
                    self.jobs[jid]["ok"] += 1
                    log("   完成 %.1f 秒 -> %s" % (r["duration"], os.path.basename(outfile)))
                else:
                    log("   失败：%s" % r["error"])
            except Exception as e:
                log("   异常：%s: %s" % (type(e).__name__, e))
                if os.environ.get("VTS2MP3_DEBUG"):
                    log(traceback.format_exc())
            finally:
                self.jobs[jid]["done"] += 1

        def runner():
            lock = threading.Semaphore(jobs)
            threads = []
            for t in targets:
                def go(t=t):
                    with lock:
                        work(t)
                th = threading.Thread(target=go, daemon=True)
                th.start()
                threads.append(th)
            for th in threads:
                th.join()
            self.jobs[jid]["finished"] = True

        threading.Thread(target=runner, daemon=True).start()
        return {"id": jid, "total": len(targets)}


def find_free_port(preferred=8756):
    import socket
    for p in [preferred] + list(range(preferred + 1, preferred + 40)):
        with socket.socket() as s:
            try:
                s.bind(("127.0.0.1", p))
                return p
            except OSError:
                continue
    return 0


def serve(args):
    port = args.port or find_free_port()
    httpd = App(("127.0.0.1", port), tmpdir=args.tmpdir)
    url = "http://127.0.0.1:%d/" % port
    print("%s v%s" % (APP_NAME, VERSION))
    print("网页界面已启动： %s" % url)
    print("（按 Ctrl+C 退出）")
    if not args.no_browser:
        threading.Timer(0.6, lambda: webbrowser.open(url)).start()
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n已退出")
    finally:
        httpd.server_close()
    return 0


def main():
    ap = argparse.ArgumentParser(
        prog=APP_NAME, description="把 .vts（实为 MP4/AAC）批量转成 MP3",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="不带参数运行即启动网页界面。")
    ap.add_argument("--cli", action="store_true", help="命令行模式")
    ap.add_argument("inputs", nargs="*", help="目录或 .vts 文件")
    ap.add_argument("-o", "--output", help="输出目录（默认 <输入目录>/mp3）")
    ap.add_argument("--src", nargs="*", help="原始录屏所在目录（可多个）")
    ap.add_argument("--bitrate", type=int, default=256, help="MP3 比特率 kbps（默认 256）")
    ap.add_argument("--port", type=int, default=0, help="网页端口（默认自动）")
    ap.add_argument("--tmpdir", help="临时目录（默认自动选择）")
    ap.add_argument("--no-browser", action="store_true", help="不自动打开浏览器")
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args()

    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

    if args.cli or args.inputs:
        if not args.inputs:
            ap.error("命令行模式需要指定目录或文件")
        return cli(args)
    return serve(args)


if __name__ == "__main__":
    sys.exit(main())
