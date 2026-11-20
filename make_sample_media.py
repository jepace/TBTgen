"""Creates placeholder audio/video clips for the sample bank (needs ffmpeg).
Real clips live in Google Drive (TBTgen/assets/media); copy them into assets/media to use them."""
import subprocess
from pathlib import Path

MEDIA = Path(__file__).resolve().parent / "assets" / "media"
FONT = next((f for f in ["/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
                          "/usr/local/share/fonts/dejavu/DejaVuSans-Bold.ttf",
                          "/System/Library/Fonts/Supplemental/Arial Bold.ttf"] if Path(f).exists()), "")
MEDIA.mkdir(parents=True, exist_ok=True)


def run(*args):
    subprocess.run(["ffmpeg", "-v", "error", "-y", *args], check=True)


for name in ["connected", "alices_restaurant", "thank_you_dido", "gravy", "thank_you_sly"]:
    out = MEDIA / f"{name}.mp3"
    if not out.exists():
        run("-f", "lavfi", "-i", "sine=frequency=440:duration=20", "-af", "volume=0.2", str(out))
        print("made", out.name)

for name in ["addams_family_values", "free_birds", "pieces_of_april", "scent_of_a_woman", "hannah_and_her_sisters"]:
    out = MEDIA / f"{name}.mp4"
    if not out.exists():
        text = "drawtext=" + (f"fontfile={FONT}:" if FONT else "") + \
               "text='PLACEHOLDER MOVIE CLIP':fontsize=64:fontcolor=white:box=1:boxcolor=black@0.7:x=(w-tw)/2:y=(h-th)/2"
        run("-f", "lavfi", "-i", "testsrc2=size=1280x720:rate=24:duration=15", "-f", "lavfi",
            "-i", "sine=frequency=330:duration=15", "-vf", text, "-c:v", "libx264", "-crf", "32",
            "-pix_fmt", "yuv420p", "-c:a", "aac", "-af", "volume=0.2", "-shortest", str(out))
        print("made", out.name)
