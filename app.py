import os
import re
import csv
import io
import zipfile
from flask import Flask, render_template, request, send_file, jsonify
from youtube_transcript_api import YouTubeTranscriptApi
from openpyxl import Workbook

app = Flask(__name__)


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


def format_timestamp(seconds):
    h = int(seconds // 3600)
    m = int((seconds % 3600) // 60)
    s = int(seconds % 60)
    ms = int((seconds - int(seconds)) * 1000)
    return f"{h:02d}:{m:02d}:{s:02d},{ms:03d}"


def fetch_transcript(video_id, max_minutes=None):
    ytt_api = YouTubeTranscriptApi()
    transcript = ytt_api.fetch(video_id)
    entries = []
    max_seconds = max_minutes * 60 if max_minutes else None
    for snippet in transcript:
        if max_seconds and snippet.start >= max_seconds:
            break
        entries.append({
            'start': snippet.start,
            'duration': snippet.duration,
            'text': snippet.text,
        })
    return entries


def build_srt(entries):
    lines = []
    for i, e in enumerate(entries, 1):
        start = format_timestamp(e['start'])
        end = format_timestamp(e['start'] + e['duration'])
        lines.append(f"{i}\n{start} --> {end}\n{e['text']}\n")
    return "\n".join(lines)


def build_csv(entries):
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(['Index', 'Start', 'End', 'Duration (s)', 'Text'])
    for i, e in enumerate(entries, 1):
        writer.writerow([
            i,
            format_timestamp(e['start']),
            format_timestamp(e['start'] + e['duration']),
            round(e['duration'], 2),
            e['text'],
        ])
    return output.getvalue()


def build_excel(entries):
    wb = Workbook()
    ws = wb.active
    ws.title = "Subtitles"
    ws.append(['Index', 'Start', 'End', 'Duration (s)', 'Text'])
    for i, e in enumerate(entries, 1):
        ws.append([
            i,
            format_timestamp(e['start']),
            format_timestamp(e['start'] + e['duration']),
            round(e['duration'], 2),
            e['text'],
        ])
    for col in ['A', 'B', 'C', 'D']:
        ws.column_dimensions[col].width = 18
    ws.column_dimensions['E'].width = 80
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/extract', methods=['POST'])
def extract():
    data = request.json
    url = data.get('url', '').strip()
    max_minutes = data.get('max_minutes')

    video_id = extract_video_id(url)
    if not video_id:
        return jsonify({'error': 'Invalid YouTube URL'}), 400

    try:
        if max_minutes:
            max_minutes = float(max_minutes)
        else:
            max_minutes = None
        entries = fetch_transcript(video_id, max_minutes)
    except Exception as e:
        return jsonify({'error': f'Could not fetch subtitles: {str(e)}'}), 400

    if not entries:
        return jsonify({'error': 'No subtitles found for this video in the specified range'}), 404

    duration_label = f"{int(max_minutes)}min" if max_minutes else "full"
    last_entry = entries[-1]
    actual_end = format_timestamp(last_entry['start'] + last_entry['duration'])

    return jsonify({
        'video_id': video_id,
        'count': len(entries),
        'duration_label': duration_label,
        'actual_end': actual_end,
    })


@app.route('/download', methods=['POST'])
def download():
    data = request.json
    url = data.get('url', '').strip()
    max_minutes = data.get('max_minutes')
    fmt = data.get('format', 'all')

    video_id = extract_video_id(url)
    if not video_id:
        return jsonify({'error': 'Invalid YouTube URL'}), 400

    if max_minutes:
        max_minutes = float(max_minutes)
    else:
        max_minutes = None

    entries = fetch_transcript(video_id, max_minutes)
    duration_label = f"{int(max_minutes)}min" if max_minutes else "full"
    base = f"{video_id}_{duration_label}"

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


if __name__ == '__main__':
    app.run(debug=True, port=5555)
