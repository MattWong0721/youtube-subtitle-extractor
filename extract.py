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


SENTENCE_ENDINGS = re.compile(r'[.!?。！？]$')
BUFFER_SECONDS = 30


def fetch_transcript(video_id, max_minutes=None):
    ytt_api = YouTubeTranscriptApi()
    transcript = ytt_api.fetch(video_id)

    all_entries = []
    for snippet in transcript:
        all_entries.append({
            'start': snippet.start,
            'duration': snippet.duration,
            'text': snippet.text,
        })

    if not max_minutes:
        return all_entries

    max_seconds = max_minutes * 60
    buffer_limit = max_seconds + BUFFER_SECONDS

    hard_cut = []
    buffer_zone = []
    for e in all_entries:
        if e['start'] < max_seconds:
            hard_cut.append(e)
        elif e['start'] < buffer_limit:
            buffer_zone.append(e)
        else:
            break

    if not buffer_zone:
        return hard_cut

    last_sentence_idx = None
    for i, e in enumerate(buffer_zone):
        if SENTENCE_ENDINGS.search(e['text'].strip()):
            last_sentence_idx = i

    if last_sentence_idx is not None:
        return hard_cut + buffer_zone[:last_sentence_idx + 1]

    return hard_cut + buffer_zone


def clamp_entries(entries):
    clamped = []
    for idx, e in enumerate(entries):
        end = e['start'] + e['duration']
        if idx + 1 < len(entries):
            next_start = entries[idx + 1]['start']
            if end > next_start:
                end = next_start
        clamped.append({**e, 'end': end})
    return clamped


def save_srt(entries, path):
    entries = clamp_entries(entries)
    with open(path, 'w', encoding='utf-8') as f:
        for i, e in enumerate(entries, 1):
            start = format_timestamp(e['start'])
            end = format_timestamp(e['end'])
            f.write(f"{i}\n{start} --> {end}\n{e['text']}\n\n")


def save_csv(entries, path):
    entries = clamp_entries(entries)
    with open(path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(['Index', 'Start', 'End', 'Duration (s)', 'Text'])
        for i, e in enumerate(entries, 1):
            writer.writerow([
                i,
                format_timestamp(e['start']),
                format_timestamp(e['end']),
                round(e['end'] - e['start'], 2),
                e['text'],
            ])


def save_excel(entries, path):
    entries = clamp_entries(entries)
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
