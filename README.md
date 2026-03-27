# Vegas Pro 22 - Whisper Auto Subtitles

## What this script does
- Uses local OpenAI Whisper CLI (`whisper`) to transcribe a selected media file.
- Reads generated `.srt` subtitles.
- Creates a new video track named `Whisper Subtitles`.
- Adds one text event per subtitle entry.

## Required files / dependencies
You already have:
- `ScriptPortal.Vegas.dll`

You also need locally:
- Python 3
- OpenAI Whisper CLI installed: `pip install -U openai-whisper`
- `ffmpeg` available in PATH

Optional:
- Set env var `WHISPER_EXE` to your whisper command/path.
  - Example: `setx WHISPER_EXE "C:\\Users\\you\\AppData\\Roaming\\Python\\Python311\\Scripts\\whisper.exe"`

## Install script in VEGAS
1. Copy `WhisperAutoSubtitles.cs` into your VEGAS Script Menu folder.
2. In VEGAS Pro 22, run the script from `Tools -> Scripting`.

## Usage
1. Select a timeline event with media (optional).
2. Run script.
3. Choose model/language and Whisper executable when prompted.
4. Wait for transcription to finish.

## Notes
- If no event is selected, the script asks you to pick a media/audio file.
- The script targets OpenAI Whisper Python CLI arguments.
- Subtitle style depends on your `Titles & Text` generator defaults.
