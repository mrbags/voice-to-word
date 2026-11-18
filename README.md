# Voice-to-Word Transcriber (Windows)

Listens to a computer voice — text-to-speech, a screen reader, narration in a video — and types what it says into a Word document. Runs fully offline after a one-time model download.

## Three ways to listen
- **This computer's audio** – captures whatever is playing through your speakers or headphones. No cables, drivers or "Stereo Mix" needed. (The voice is captured even if your volume is low, but not if the app playing it is muted.)
- **Microphone** – for a voice playing on another device in the room.
- **Audio file** – transcribes an .mp3, .wav, .m4a, .mp4 and similar, with a progress bar.

## Getting it running
You need **Python 3.9 or newer** from python.org (tick "Add Python to PATH" when installing).

- **Quick start:** double-click `run.bat`. The first run installs what it needs.
- **Standalone app:** double-click `build_exe.bat`. After a few minutes you'll have `dist\VoiceToWord\VoiceToWord.exe` — copy that folder to any Windows PC, no Python needed there.

The first time you press Start, the app downloads the speech model you picked (about 75 MB for *base.en*, 250 MB for *small.en*, 770 MB for *medium.en*) into a `models` folder beside the app. After that it works without internet.

## Using it
1. Pick the source and device (your usual speakers are pre-selected).
2. Pick an accuracy level. *small.en* is the best balance on most PCs; use *base.en* on older laptops if text falls behind the voice.
3. Choose where to save, then press **Start listening** and play the voice.
4. Press **Stop** when done, then **Open document**.

Text appears live in the window and the .docx is saved after every sentence, so nothing is lost if the PC shuts down. A pause of more than 2 seconds starts a new paragraph. Tick *Add timestamps* to prefix each paragraph with the time it started.

If the document is open in Word while the app is writing, Word locks the file — the app warns you and catches up once you close it.

## Tips
- Computer voices are very clean, so even *base.en* is usually near-perfect. Use *large-v3* for non-English voices.
- For a very long recording, saving it to a file first and using **Audio file** mode is faster than real time.

## How the code is organised
Everything lives in `voice_to_word.py`, which is commented throughout. In order:
- **DocWriter** builds the Word document and saves it after every sentence.
- **Segmenter** splits live audio into sentences by detecting pauses.
- **AudioCapture** and **list_devices** read audio from Windows sound devices.
- **App** is the window. Its background threads capture audio and run the Whisper model.

## Author
Created by Danny M. Byers, [Native Thought](https://www.nativethought.com).

## License
Released under the [MIT License](LICENSE). Speech recognition uses OpenAI's Whisper models (MIT) through [faster-whisper](https://github.com/SYSTRAN/faster-whisper) (MIT).
