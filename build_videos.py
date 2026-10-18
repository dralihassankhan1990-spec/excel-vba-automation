"""Assemble the two narrated demo videos (TTS + Ken Burns stills + concat).

Stills are REAL outputs: Task Export / Dashboard renders, tracker render,
Weekly Report Summary/Overdue renders, and terminal cards quoting real log lines.
Usage: python3 build_videos.py
"""
import os
import subprocess

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(HERE, "video_assets")
SEG = os.path.join(HERE, "video_segments")
AUD = os.path.join(HERE, "video_audio")
os.makedirs(SEG, exist_ok=True)
os.makedirs(AUD, exist_ok=True)

TTS = "/opt/hatch/bin/tts"
FPS, VW, VH = 30, 1280, 720


def sh(*args, **kw):
    subprocess.run(args, check=True, **kw)


def probe_dur(path):
    r = subprocess.run(["ffprobe", "-v", "error", "-show_entries", "format=duration",
                        "-of", "default=noprint_wrappers=1:nokey=1", path],
                       capture_output=True, text=True, check=True)
    return float(r.stdout.strip())


def narrate(key, text):
    out = os.path.join(AUD, f"{key}.mp3")
    if not os.path.exists(out):
        sh(TTS, "speak", "--output", out, "--text-stdin", input=text, text=True)
    return out


def seg_image(key, img, text, min_dur):
    audio = narrate(key, text)
    dur = max(probe_dur(audio) + 0.8, min_dur)
    out = os.path.join(SEG, f"{key}.mp4")
    if os.path.exists(out):
        return out
    frames = int(dur * FPS)
    vf = (f"scale={VW}:{VH}:force_original_aspect_ratio=decrease,"
          f"pad={VW}:{VH}:(ow-iw)/2:(oh-ih)/2,"
          f"zoompan=z='min(zoom+0.0012,1.12)':d={frames}:"
          f"x='iw/2-(iw/zoom/2)':y='ih/2-(ih/zoom/2)':s={VW}x{VH}:fps={FPS}")
    sh("ffmpeg", "-y", "-v", "error", "-loop", "1", "-i", img, "-i", audio,
       "-vf", vf, "-af", "apad", "-t", f"{dur:.2f}",
       "-c:v", "libx264", "-pix_fmt", "yuv420p", "-r", str(FPS),
       "-c:a", "aac", "-shortest", out)
    return out


def build(name, segments, final):
    parts = []
    for key, img, text, min_dur in segments:
        parts.append(seg_image(key, img, text, min_dur))
        print("built segment", key, flush=True)
    lst = os.path.join(SEG, f"{name}-concat.txt")
    with open(lst, "w") as f:
        for p in parts:
            f.write(f"file '{p}'\n")
    sh("ffmpeg", "-y", "-v", "error", "-f", "concat", "-safe", "0",
       "-i", lst, "-c", "copy", final)
    print(f"FINAL: {final} ({probe_dur(final):.1f}s)")


V1_BEFORE = os.path.join(ASSETS, "v1_problem_crop.png")
V1_AFTER = os.path.join(ASSETS, "v1_after_crop.png")
V2_BEFORE = os.path.join(ASSETS, "v2_problem_crop.png")
V2_SUMMARY = os.path.join(ASSETS, "v2_summary_crop.png")
V2_OVERDUE = os.path.join(ASSETS, "v2_overdue_crop.png")

V1 = [
    ("v1_title", os.path.join(ASSETS, "v1_title.png"),
     "Task Dashboard. One click turns a flat Excel task export into a live project dashboard. Built with V B A, by Muhammad Bilal.", 6),
    ("v1_problem", V1_BEFORE,
     "Here is the problem. A project team tracks forty tasks in a plain export: titles, owners, statuses, due dates. Every week, someone spends half an hour counting what is overdue, who is behind, and what is due next, by hand. The numbers change between meetings, and nobody trusts them.", 10),
    ("v1_how", os.path.join(ASSETS, "v1_how.png"),
     "The solution is a single V B A macro called Build Dashboard. It reads the Task Export sheet, counts the K P Is, aggregates tasks by owner and by status, and rebuilds a formatted Dashboard sheet: K P I tiles, per-owner completion rates, and two action lists — overdue in red, due in the next seven days in amber. Re-run it any time and the Dashboard rebuilds from the latest data.", 15),
    ("v1_run", os.path.join(ASSETS, "v1_terminal.png"),
     "Here it is running for real. The macro executes, the embedded V B A project is verified byte-identical to the source, and the Dashboard checks pass: forty tasks, twelve completed, fourteen in progress, nine overdue, thirty percent complete.", 12),
    ("v1_after", V1_AFTER,
     "And here is the result: a formatted Dashboard. Five K P I tiles at the top, tasks by owner and by status, nine overdue tasks flagged in red, seven due this week flagged in amber, all from one click.", 12),
    ("v1_result", os.path.join(ASSETS, "v1_result.png"),
     "Forty tasks, one click, live dashboard. No manual counting, no stale numbers. Fully local — no network, no A P I keys.", 8),
]

V2 = [
    ("v2_title", os.path.join(ASSETS, "v2_title.png"),
     "Weekly Report Generator. The Monday status pack, built automatically. Python and openpyxl, by Muhammad Bilal.", 6),
    ("v2_problem", V2_BEFORE,
     "The problem: every Monday morning, the team hand-builds the same status pack — count tasks, find what is overdue, list what is due this week, tally workload per owner — then format it and export a P D F. Thirty to forty-five minutes, error-prone, and never on time.", 10),
    ("v2_how", os.path.join(ASSETS, "v2_how.png"),
     "The solution is one script: weekly report dot py. It reads the Task Tracker, computes the K P Is, and writes a four-sheet Excel report — Summary, Overdue, Due This Week, and Workload by Owner — plus a P D F, ready to share. A scheduler runs it every Monday at eight A M, on Windows Task Scheduler or cron. Re-runs overwrite the same dated files, so there are never duplicates.", 15),
    ("v2_run", os.path.join(ASSETS, "v2_terminal.png"),
     "Running it for real: thirty tasks loaded, the Excel report written, the P D F converted — eight overdue, seven due this week, twenty percent complete. Same command, same result, every week.", 12),
    ("v2_summary", V2_SUMMARY,
     "Here is the output: the Summary sheet with the K P I table — thirty tasks, six completed, eight overdue, seven due this week.", 10),
    ("v2_overdue", V2_OVERDUE,
     "And the Overdue sheet: eight tasks in red, sorted oldest first, with days overdue. The P D F is identical, ready to send.", 10),
    ("v2_result", os.path.join(ASSETS, "v2_result.png"),
     "Thirty tasks, one command, Excel plus P D F — every Monday at eight, automatically. Fully local, no network, no A P I keys.", 8),
]


def main():
    build("v1", V1, os.path.join(HERE, "task-dashboard-vba", "task-dashboard-vba-demo.mp4"))
    build("v2", V2, os.path.join(HERE, "weekly-report-generator", "weekly-report-generator-demo.mp4"))


if __name__ == "__main__":
    main()
