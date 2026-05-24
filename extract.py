import sys
import re
import csv
import os
from youtube_transcript_api import YouTubeTranscriptApi
from openpyxl import Workbook


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
        start = snippet.start
        if max_seconds and start >= max_seconds:
            break
        entries.append({
            'start': snippet.start,
            'duration': snippet.duration,
            'text': snippet.text,
        })
    return entries


def save_srt(entries, path):
    with open(path, 'w', encoding='utf-8') as f:
        for i, e in enumerate(entries, 1):
            start = format_timestamp(e['start'])
            end = format_timestamp(e['start'] + e['duration'])
            f.write(f"{i}\n{start} --> {end}\n{e['text']}\n\n")


def save_csv(entries, path):
    with open(path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(['Index', 'Start', 'End', 'Duration (s)', 'Text'])
        for i, e in enumerate(entries, 1):
            writer.writerow([
                i,
                format_timestamp(e['start']),
                format_timestamp(e['start'] + e['duration']),
                round(e['duration'], 2),
                e['text'],
            ])


def save_excel(entries, path):
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
    for col in ['A', 'B', 'C', 'D', 'E']:
        ws.column_dimensions[col].width = 20 if col != 'E' else 80
    wb.save(path)


def main():
    url = input("Paste YouTube URL: ").strip() if len(sys.argv) < 2 else sys.argv[1]
    video_id = extract_video_id(url)
    if not video_id:
        print("Could not extract video ID from URL.")
        sys.exit(1)

    max_min_input = input("Max duration in minutes (leave blank for full video): ").strip() if len(sys.argv) < 3 else sys.argv[2]
    max_minutes = float(max_min_input) if max_min_input else None

    print(f"Fetching transcript for video {video_id}...")
    entries = fetch_transcript(video_id, max_minutes)
    print(f"Got {len(entries)} subtitle entries.")

    out_dir = os.path.dirname(os.path.abspath(__file__))
    base = f"{video_id}_{int(max_minutes)}min" if max_minutes else video_id

    srt_path = os.path.join(out_dir, f"{base}.srt")
    csv_path = os.path.join(out_dir, f"{base}.csv")
    xlsx_path = os.path.join(out_dir, f"{base}.xlsx")

    save_srt(entries, srt_path)
    save_csv(entries, csv_path)
    save_excel(entries, xlsx_path)

    print(f"\nSaved:\n  SRT:   {srt_path}\n  CSV:   {csv_path}\n  Excel: {xlsx_path}")


if __name__ == "__main__":
    main()
