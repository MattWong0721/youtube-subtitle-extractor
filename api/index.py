import re
import csv
import io
import zipfile
from flask import Flask, request, send_file, jsonify
from youtube_transcript_api import YouTubeTranscriptApi
from openpyxl import Workbook

app = Flask(__name__)

SENTENCE_ENDINGS = re.compile(r'[.!?。！？]$')
PUNCTUATION_BREAK = re.compile(r'[.!?,;:。！？，；：]$')
MAX_MERGE_SECONDS = 7
BUFFER_SECONDS = 30


def extract_video_id(url):
    patterns = [
        r'(?:v=)([a-zA-Z0-9_-]{11})',
        r'(?:youtu\.be/)([a-zA-Z0-9_-]{11})',
    ]
    for p in patterns:
        m = re.search(p, url)
        if m:
            return m.group(1)
    return None


def parse_time(time_str):
    if not time_str or not time_str.strip():
        return None
    parts = time_str.strip().replace('.', ':').split(':')
    parts = [p for p in parts if p]
    nums = [int(p) for p in parts]
    if len(nums) == 1:
        return nums[0] * 60
    elif len(nums) == 2:
        return nums[0] * 60 + nums[1]
    elif len(nums) >= 3:
        return nums[0] * 3600 + nums[1] * 60 + nums[2]
    return None


def format_timestamp(seconds):
    h = int(seconds // 3600)
    m = int((seconds % 3600) // 60)
    s = int(seconds % 60)
    ms = int((seconds - int(seconds)) * 1000)
    return f"{h:02d}:{m:02d}:{s:02d},{ms:03d}"


def format_time_short(seconds):
    h = int(seconds // 3600)
    m = int((seconds % 3600) // 60)
    s = int(seconds % 60)
    if h > 0:
        return f"{h}h{m:02d}m{s:02d}s"
    return f"{m}m{s:02d}s"


def fetch_transcript(video_id, start_sec=None, end_sec=None, smart_cut=True):
    ytt_api = YouTubeTranscriptApi()
    transcript = ytt_api.fetch(video_id)

    all_entries = []
    for snippet in transcript:
        all_entries.append({
            'start': snippet.start,
            'duration': snippet.duration,
            'text': snippet.text,
        })

    if start_sec is None:
        start_sec = 0

    filtered = [e for e in all_entries if e['start'] >= start_sec]

    if end_sec is None:
        return filtered

    hard_cut = [e for e in filtered if e['start'] < end_sec]

    if not smart_cut:
        return hard_cut

    buffer_limit = end_sec + BUFFER_SECONDS
    buffer_zone = [e for e in filtered if end_sec <= e['start'] < buffer_limit]

    if not buffer_zone:
        return hard_cut

    last_sentence_idx = None
    for i, e in enumerate(buffer_zone):
        if SENTENCE_ENDINGS.search(e['text'].strip()):
            last_sentence_idx = i

    if last_sentence_idx is not None:
        return hard_cut + buffer_zone[:last_sentence_idx + 1]

    return hard_cut + buffer_zone


def merge_subtitles(entries):
    if not entries:
        return []
    merged = []
    buf_texts = []
    buf_start = entries[0]['start']

    for idx, e in enumerate(entries):
        buf_texts.append(e['text'].strip())
        end = e['start'] + e['duration']
        if idx + 1 < len(entries):
            next_start = entries[idx + 1]['start']
            if end > next_start:
                end = next_start

        text_so_far = ' '.join(buf_texts)
        has_punctuation = PUNCTUATION_BREAK.search(text_so_far)
        too_long = (end - buf_start) >= MAX_MERGE_SECONDS
        is_last = idx == len(entries) - 1

        if has_punctuation or too_long or is_last:
            merged.append({
                'start': buf_start,
                'end': end,
                'text': text_so_far,
            })
            if not is_last:
                buf_texts = []
                buf_start = entries[idx + 1]['start']

    return merged


def build_srt(entries):
    entries = merge_subtitles(entries)
    lines = []
    for i, e in enumerate(entries, 1):
        start = format_timestamp(e['start'])
        end = format_timestamp(e['end'])
        lines.append(f"{i}\n{start} --> {end}\n{e['text']}\n")
    return "\n".join(lines)


def build_csv(entries):
    entries = merge_subtitles(entries)
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(['Index', 'Start', 'End', 'Duration (s)', 'Text'])
    for i, e in enumerate(entries, 1):
        writer.writerow([
            i,
            format_timestamp(e['start']),
            format_timestamp(e['end']),
            round(e['end'] - e['start'], 2),
            e['text'],
        ])
    return output.getvalue()


def build_excel(entries):
    entries = merge_subtitles(entries)
    wb = Workbook()
    ws = wb.active
    ws.title = "Subtitles"
    ws.append(['Index', 'Start', 'End', 'Duration (s)', 'Text'])
    for i, e in enumerate(entries, 1):
        ws.append([
            i,
            format_timestamp(e['start']),
            format_timestamp(e['end']),
            round(e['end'] - e['start'], 2),
            e['text'],
        ])
    for col in ['A', 'B', 'C', 'D']:
        ws.column_dimensions[col].width = 18
    ws.column_dimensions['E'].width = 80
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


def make_label(start_sec, end_sec):
    parts = []
    if start_sec:
        parts.append(format_time_short(start_sec))
    else:
        parts.append("0m00s")
    parts.append(format_time_short(end_sec) if end_sec else "end")
    return "-".join(parts)


@app.route('/api/extract', methods=['POST'])
def extract():
    data = request.json
    url = data.get('url', '').strip()
    start_time = data.get('start_time')
    end_time = data.get('end_time')
    smart_cut = data.get('smart_cut', True)

    video_id = extract_video_id(url)
    if not video_id:
        return jsonify({'error': 'Invalid YouTube URL'}), 400

    try:
        start_sec = parse_time(start_time)
        end_sec = parse_time(end_time)
        entries = fetch_transcript(video_id, start_sec, end_sec, smart_cut)
    except Exception as e:
        return jsonify({'error': f'Could not fetch subtitles: {str(e)}'}), 400

    if not entries:
        return jsonify({'error': 'No subtitles found for this video in the specified range'}), 404

    first_entry = entries[0]
    last_entry = entries[-1]
    actual_start = format_timestamp(first_entry['start'])
    actual_end_sec = last_entry['start'] + last_entry['duration']
    actual_end = format_timestamp(actual_end_sec)

    extended = False
    if smart_cut and end_sec and actual_end_sec > end_sec:
        extended = True

    return jsonify({
        'video_id': video_id,
        'count': len(entries),
        'actual_start': actual_start,
        'actual_end': actual_end,
        'extended': extended,
    })


@app.route('/api/download', methods=['POST'])
def download():
    data = request.json
    url = data.get('url', '').strip()
    start_time = data.get('start_time')
    end_time = data.get('end_time')
    smart_cut = data.get('smart_cut', True)
    fmt = data.get('format', 'all')

    video_id = extract_video_id(url)
    if not video_id:
        return jsonify({'error': 'Invalid YouTube URL'}), 400

    start_sec = parse_time(start_time)
    end_sec = parse_time(end_time)
    entries = fetch_transcript(video_id, start_sec, end_sec, smart_cut)

    label = make_label(start_sec, end_sec)
    base = f"{video_id}_{label}"

    if fmt == 'srt':
        content = build_srt(entries)
        return send_file(
            io.BytesIO(content.encode('utf-8')),
            mimetype='text/plain',
            as_attachment=True,
            download_name=f"{base}.srt",
        )
    elif fmt == 'csv':
        content = build_csv(entries)
        return send_file(
            io.BytesIO(content.encode('utf-8')),
            mimetype='text/csv',
            as_attachment=True,
            download_name=f"{base}.csv",
        )
    elif fmt == 'excel':
        content = build_excel(entries)
        return send_file(
            io.BytesIO(content),
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f"{base}.xlsx",
        )
    else:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
            zf.writestr(f"{base}.srt", build_srt(entries))
            zf.writestr(f"{base}.csv", build_csv(entries))
            zf.writestr(f"{base}.xlsx", build_excel(entries))
        buf.seek(0)
        return send_file(
            buf,
            mimetype='application/zip',
            as_attachment=True,
            download_name=f"{base}.zip",
        )
