"""
Voice-to-Word Transcriber
=========================

A small Windows desktop app that listens to a computer voice (text-to-speech,
a screen reader, narration in a video, ...) and types what it says into a
Microsoft Word (.docx) document. Everything runs offline on your own PC.

Where the audio can come from
-----------------------------
  * This computer's audio - whatever is playing through the speakers or
                            headphones. Uses Windows "WASAPI loopback", so no
                            extra drivers or "Stereo Mix" are needed.
  * Microphone            - a voice playing on another device in the room.
  * Audio file            - a saved recording (mp3, wav, m4a, mp4, ...).

How it works, in one paragraph
------------------------------
For live audio, small chunks of sound arrive from Windows about 20 times a
second. A "Segmenter" watches how loud each chunk is and cuts the stream into
utterances wherever the voice pauses. Each utterance is handed to a background
thread that runs OpenAI's Whisper speech-recognition model (through the
faster-whisper library) and turns it into text. The text is shown in the
window and appended to the Word document, which is saved after every
utterance so nothing is lost if the app or PC stops unexpectedly. Audio files
skip the Segmenter: Whisper reads the whole file and returns its sentences.

Whisper spells unfamiliar names the way they sound ("Kehind Wiley"). The
optional "Words to expect" box lets the user list names and terms in advance;
they are passed to Whisper as hints so it spells them correctly.

Author: Danny M. Byers, Native Thought (https://www.nativethought.com)
License: MIT (see the LICENSE file)

Code layout (top to bottom)
---------------------------
  1. Imports and settings
  2. Small helper functions
  3. DocWriter      - builds and saves the Word document
  4. Audio helpers  - convert raw audio into the format Whisper expects
  5. Segmenter      - splits live audio into utterances at pauses
  6. AudioCapture   - reads audio from a Windows sound device
  7. list_devices   - finds the speakers and microphones on this PC
  8. App            - the window (Tkinter GUI) and the worker threads
  9. main()         - program entry point

Threads, briefly
----------------
Tkinter (the GUI toolkit) is not thread-safe: only the main thread may touch
widgets. Background threads therefore never update the window directly.
Instead they call App.post(), which puts a small "please run this function"
message on a queue, and the main thread drains that queue every 100 ms.
"""

# --------------------------------------------------------------------------- #
#  1. Imports and settings
# --------------------------------------------------------------------------- #

# Python standard library
import os                 # file paths and folders
import sys                # detecting Windows / a packaged .exe
import queue              # thread-safe queues for passing work between threads
import threading          # running audio capture and transcription off the GUI thread
import datetime           # dates for file names and the document subtitle
import traceback          # printing full error details to the console
from collections import deque   # fixed-size buffer for the "pre-roll" audio

# Tkinter: the GUI toolkit that ships with Python on Windows
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

# Third-party packages (installed from requirements.txt)
import numpy as np                 # fast maths on arrays of audio samples
from docx import Document          # python-docx: creates .docx files
from docx.shared import Pt         # font sizes in points

# PyAudioWPatch is a Windows-only fork of PyAudio that adds "loopback"
# recording, i.e. capturing what the speakers are playing. On other systems it
# is not installed; we set it to None so the rest of the app (audio-file mode)
# still works and the live options simply explain what's missing.
try:
    import pyaudiowpatch as pyaudio
except ImportError:
    pyaudio = None

# The name shown in the title bar and in pop-up dialogs.
APP_NAME = "Voice-to-Word Transcriber"

# Whisper models are trained on 16 kHz mono audio, so every live recording is
# converted to this rate before transcription.
TARGET_SR = 16000

# Whisper model sizes offered in the "Accuracy" drop-down. Names ending in
# ".en" are English-only models (smaller and more accurate for English);
# "large-v3" understands many languages. Each one is downloaded on first use.
MODELS = ["base.en", "small.en", "medium.en", "tiny.en", "large-v3"]

# A short hint shown next to the drop-down for whichever model is selected.
MODEL_NOTES = {
    "tiny.en": "fastest, least accurate",
    "base.en": "fast, good for clear computer voices",
    "small.en": "balanced (recommended)",
    "medium.en": "slow, very accurate",
    "large-v3": "slowest, any language",
}

# If the voice is silent for longer than this many seconds, the next words
# start a new paragraph in the Word document.
PARAGRAPH_GAP_S = 2.0

# The "Words to expect" list is remembered between runs in this small text
# file, stored beside the app (see app_dir()).
WORDS_FILE_NAME = "words.txt"


# --------------------------------------------------------------------------- #
#  2. Small helper functions
# --------------------------------------------------------------------------- #

def app_dir():
    """Return the folder the app lives in.

    When run as a script, that's the folder containing this .py file. When
    packaged by PyInstaller into VoiceToWord.exe, `sys.frozen` is set and the
    folder is the one containing the .exe. Downloaded speech models are stored
    in a "models" subfolder here so they travel with the app.
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def default_output_folder():
    """The folder transcripts are saved to by default: the user's Documents
    folder, or their home folder if Documents doesn't exist."""
    docs = os.path.join(os.path.expanduser("~"), "Documents")
    return docs if os.path.isdir(docs) else os.path.expanduser("~")


def timestamped_name():
    """A unique file name such as 'Transcript_2026-11-18_120455.docx', so a
    new session never overwrites an earlier transcript."""
    return "Transcript_" + datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S") + ".docx"


def fmt_clock(seconds):
    """Format a number of seconds as 'mm:ss', or 'h:mm:ss' past one hour.
    Used for the optional paragraph timestamps and the progress text."""
    seconds = int(seconds)
    h, rem = divmod(seconds, 3600)
    m, s = divmod(rem, 60)
    return f"{h}:{m:02d}:{s:02d}" if h else f"{m:02d}:{s:02d}"


# --------------------------------------------------------------------------- #
#  3. DocWriter - builds and saves the Word document
# --------------------------------------------------------------------------- #

class DocWriter:
    """Builds the Word document piece by piece as text arrives.

    The document is saved to disk after every addition, so if the app is
    closed or the computer shuts down mid-session, everything transcribed up
    to that moment is already in the file.
    """

    def __init__(self, path, title, timestamps):
        """Create a new document with a heading and a 'Transcribed on ...'
        subtitle, and save it straight away.

        path       - where to save the .docx
        title      - the heading at the top of the document
        timestamps - if True, each paragraph starts with a bold [mm:ss] time
        """
        self.path = path
        self.timestamps = timestamps
        self.doc = Document()

        # Set the default font for body text.
        normal = self.doc.styles["Normal"]
        normal.font.name = "Calibri"
        normal.font.size = Pt(11)

        # Heading, then an italic line recording when the transcript was made.
        self.doc.add_heading(title, level=1)
        sub = self.doc.add_paragraph()
        run = sub.add_run("Transcribed " + datetime.datetime.now().strftime("%B %d, %Y at %I:%M %p"))
        run.italic = True

        # The paragraph currently being written to. None means "no body text
        # yet", so the first call to add() always starts a paragraph.
        self.para = None
        self.save()

    def add(self, text, new_paragraph, stamp=None):
        """Append transcribed text to the document and save it.

        text          - the words to add
        new_paragraph - True to start a new paragraph, False to continue the
                        current one (joined with a space)
        stamp         - optional 'mm:ss' label used when timestamps are on

        Returns True if the save succeeded, False if the file is locked
        (usually because it's open in Word). The text stays in memory either
        way and is written on the next successful save.
        """
        text = text.strip()
        if not text:
            return True  # nothing to add, nothing to save
        if self.para is None or new_paragraph:
            self.para = self.doc.add_paragraph()
            if self.timestamps and stamp:
                self.para.add_run(f"[{stamp}] ").bold = True
            self.para.add_run(text)
        else:
            self.para.add_run(" " + text)
        return self.save()

    def save(self):
        """Write the document to disk. Returns False instead of crashing when
        Windows refuses because the file is open in another program."""
        try:
            self.doc.save(self.path)
            return True
        except PermissionError:
            return False


# --------------------------------------------------------------------------- #
#  4. Audio helpers - convert raw audio into what Whisper expects
# --------------------------------------------------------------------------- #

def to_mono_float(raw, channels):
    """Convert raw audio bytes from the sound card into mono float samples.

    The sound card delivers 16-bit integers (-32768..32767), interleaved by
    channel (left, right, left, right, ...). Whisper wants single-channel
    floats between -1.0 and 1.0, so we scale the numbers and average the
    channels together.
    """
    a = np.frombuffer(raw, dtype=np.int16).astype(np.float32) / 32768.0
    if channels > 1:
        # Drop any incomplete trailing frame, reshape to (samples, channels),
        # then average across channels to get mono.
        a = a[: len(a) - len(a) % channels].reshape(-1, channels).mean(axis=1)
    return a


def resample(audio, src_rate, dst_rate=TARGET_SR):
    """Change the sample rate of `audio` from src_rate to dst_rate.

    Sound cards usually run at 44.1 or 48 kHz; Whisper needs 16 kHz. Simple
    linear interpolation is plenty for speech recognition and avoids an extra
    heavyweight dependency.
    """
    if src_rate == dst_rate or len(audio) == 0:
        return audio.astype(np.float32)
    n_out = int(round(len(audio) * dst_rate / src_rate))
    x_old = np.arange(len(audio), dtype=np.float64)   # positions of input samples
    x_new = np.linspace(0, len(audio) - 1, n_out)     # positions we want samples at
    return np.interp(x_new, x_old, audio).astype(np.float32)


# --------------------------------------------------------------------------- #
#  5. Segmenter - splits live audio into utterances at pauses
# --------------------------------------------------------------------------- #

class Segmenter:
    """Cuts a continuous live audio stream into utterances.

    Whisper works on finished pieces of audio, not on a never-ending stream,
    so we need to decide where one piece ends and the next begins. The best
    place to cut is a pause between sentences, so words are never split.

    How it decides what is speech: audio is examined in 30 ms "frames". For
    each frame we measure its loudness (RMS, root-mean-square). While nobody
    is speaking we keep a running average of the background noise level; a
    frame counts as speech when it is clearly louder than that (3x the noise,
    with a fixed minimum so near-silent systems still work).

    How it decides where to cut:
      * a pause of END_SILENCE_S ends the utterance;
      * very long stretches without a clear pause (a fast-talking voice) are
        cut at any short pause once they pass SOFT_MAX_S, and always before
        HARD_MAX_S, because Whisper only looks at 30 seconds at a time.

    Time is measured on an "audio clock" (seconds of audio received since
    starting), not the wall clock, so paragraph breaks and timestamps stay
    correct even if the computer is busy and processing falls behind.

    Each finished utterance is passed to:
        on_segment(audio_16k, new_paragraph, start_seconds)
    """

    FRAME_S = 0.03           # length of one analysis frame, in seconds
    END_SILENCE_S = 0.7      # a pause this long ends an utterance
    SOFT_MAX_S = 15.0        # past this length, cut at any short pause...
    SOFT_SILENCE_S = 0.25    # ...this long
    HARD_MAX_S = 28.0        # never let an utterance grow beyond this
    MIN_SPEECH_S = 0.3       # ignore blips shorter than this (clicks, pops)
    PREROLL_FRAMES = 10      # keep 0.3 s of audio from *before* speech starts,
                             # so the first syllable isn't clipped off

    def __init__(self, rate, on_segment):
        """rate       - sample rate of the incoming audio (e.g. 48000)
        on_segment - function to call with each finished utterance"""
        self.rate = rate
        self.on_segment = on_segment
        self.frame_len = max(1, int(rate * self.FRAME_S))   # samples per frame

        # Leftover samples that didn't fill a whole frame yet.
        self.pending = np.zeros(0, dtype=np.float32)
        # The last few quiet frames, kept as "pre-roll" (see PREROLL_FRAMES).
        self.pre = deque(maxlen=self.PREROLL_FRAMES)
        # Frames belonging to the utterance currently being collected.
        self.frames = []

        self.in_speech = False    # are we in the middle of an utterance?
        self.speech_s = 0.0       # seconds of actual speech in this utterance
        self.seg_s = 0.0          # total length of this utterance so far
        self.silence_s = 0.0      # length of the current pause
        self.noise = 0.002        # running estimate of background loudness

        self.clock = 0.0          # audio clock: seconds received so far
        self.last_end = None      # audio-clock time the previous utterance ended
        self.seg_start = 0.0      # audio-clock time this utterance started
        self.seg_new_para = True  # should this utterance start a new paragraph?
        self.silence_s_at_end = 0.0

    def feed(self, samples):
        """Add newly received mono samples and process every complete frame."""
        self.pending = np.concatenate([self.pending, samples])
        n = len(self.pending) // self.frame_len
        for i in range(n):
            self._frame(self.pending[i * self.frame_len:(i + 1) * self.frame_len])
        # Keep the remainder for next time.
        self.pending = self.pending[n * self.frame_len:]

    def feed_silence(self, seconds):
        """Account for time when the sound device sent nothing at all.

        Windows loopback sends no data while the computer is completely
        silent. Without this, a pause at the end of a sentence would never be
        "heard" and the last utterance would wait forever. During speech we
        feed real zero-valued samples so the pause logic runs; otherwise we
        just advance the clock.
        """
        if self.in_speech:
            self.feed(np.zeros(int(self.rate * seconds), dtype=np.float32))
        else:
            self.clock += seconds

    def _frame(self, f):
        """Process one 30 ms frame: decide if it is speech, and whether the
        current utterance should end here."""
        self.clock += self.FRAME_S
        rms = float(np.sqrt(np.mean(f * f)))          # loudness of this frame
        threshold = max(0.006, self.noise * 3.0)      # what counts as "speech"

        if rms > threshold:
            # ---- A loud frame: speech ----
            if not self.in_speech:
                # Speech just began. Start a new utterance, beginning with the
                # pre-roll so the first sound isn't cut off.
                self.in_speech = True
                self.frames = list(self.pre)
                self.speech_s = self.seg_s = 0.0
                self.seg_start = max(0.0, self.clock - self.FRAME_S * len(self.frames))
                # A long gap since the previous utterance means a new paragraph.
                self.seg_new_para = (self.last_end is None or
                                     self.clock - self.last_end > PARAGRAPH_GAP_S)
            self.frames.append(f)
            self.speech_s += self.FRAME_S
            self.seg_s += self.FRAME_S
            self.silence_s = 0.0

        elif not self.in_speech:
            # ---- A quiet frame between utterances ----
            # Slowly update the background-noise estimate (an exponential
            # moving average) and remember the frame as possible pre-roll.
            self.noise = 0.95 * self.noise + 0.05 * rms
            self.pre.append(f)
            return

        else:
            # ---- A quiet frame inside an utterance: a pause ----
            # Keep it (natural pauses help Whisper with punctuation) and check
            # whether the pause is long enough to end the utterance.
            self.frames.append(f)
            self.seg_s += self.FRAME_S
            self.silence_s += self.FRAME_S
            if self.silence_s >= self.END_SILENCE_S:
                self._finish(continued=False)
                return
            if self.seg_s >= self.SOFT_MAX_S and self.silence_s >= self.SOFT_SILENCE_S:
                self._finish(continued=True)
                return

        # Safety net: never exceed Whisper's 30-second window.
        if self.seg_s >= self.HARD_MAX_S:
            self._finish(continued=True)

    def _finish(self, continued):
        """Close the current utterance and hand it to on_segment.

        continued - True when we had to cut in the middle of a long stretch
                    of speech. The voice is still talking, so collection
                    restarts immediately and the next piece continues the
                    same paragraph.
        """
        self.silence_s_at_end = self.silence_s

        # Only send it on if it contained enough real speech to be worth
        # transcribing (this filters out clicks and short noises).
        if self.speech_s >= self.MIN_SPEECH_S and self.frames:
            audio = resample(np.concatenate(self.frames), self.rate)
            self.on_segment(audio, self.seg_new_para, self.seg_start)

        # Reset for the next utterance.
        self.in_speech = False
        self.frames = []
        self.pre.clear()
        self.silence_s = 0.0

        if continued:
            # Long run of speech cut mid-flow: carry on collecting straight away.
            self.in_speech = True
            self.speech_s = self.seg_s = 0.0
            self.seg_start = self.clock
            self.seg_new_para = False
        else:
            # Remember when the speech actually stopped (before the trailing
            # pause) so the next paragraph-break decision is accurate.
            self.last_end = self.clock - self.silence_s_at_end

    def flush(self):
        """Send whatever is still being collected (called when the user
        presses Stop, so the last words aren't lost)."""
        if self.in_speech:
            self._finish(continued=False)


# --------------------------------------------------------------------------- #
#  6. AudioCapture - reads audio from a Windows sound device
# --------------------------------------------------------------------------- #

class AudioCapture:
    """Opens a WASAPI sound device (a speaker's loopback or a microphone) and
    streams raw 16-bit audio into a queue.

    PyAudio calls our `callback` from its own internal thread every 50 ms with
    a new chunk of audio. The callback must return quickly, so it only puts
    the bytes on the queue; all processing happens elsewhere.
    """

    def __init__(self, device_info, out_q):
        """device_info - a device dictionary from list_devices()
        out_q       - queue that receives raw audio byte strings"""
        self.pa = pyaudio.PyAudio()
        # Record in the device's own format to avoid Windows having to convert.
        self.rate = int(device_info["defaultSampleRate"])
        self.channels = max(1, int(device_info["maxInputChannels"]))

        def callback(in_data, frame_count, time_info, status):
            out_q.put(in_data)
            return (None, pyaudio.paContinue)   # keep the stream running

        self.stream = self.pa.open(
            format=pyaudio.paInt16,                  # 16-bit samples
            channels=self.channels,
            rate=self.rate,
            input=True,                              # we are recording, not playing
            input_device_index=device_info["index"],
            frames_per_buffer=int(self.rate * 0.05), # 50 ms per callback
            stream_callback=callback,
        )
        self.stream.start_stream()

    def stop(self):
        """Stop and release the sound device. Each step is attempted even if
        an earlier one fails, so the device is never left locked."""
        for fn in (self.stream.stop_stream, self.stream.close, self.pa.terminate):
            try:
                fn()
            except Exception:
                pass


# --------------------------------------------------------------------------- #
#  7. list_devices - finds the speakers and microphones on this PC
# --------------------------------------------------------------------------- #

def list_devices(kind):
    """List the sound devices available for live capture.

    kind - "loopback" for speakers/headphones (records what they play),
           "mic" for microphones.

    Returns (devices, default_index): the list of PyAudio device
    dictionaries, and the position in that list of the Windows default
    device so it can be pre-selected (-1 if there are no devices).
    """
    if pyaudio is None:
        return [], -1
    pa = pyaudio.PyAudio()
    try:
        if kind == "loopback":
            # PyAudioWPatch exposes every output device as a "loopback"
            # input device; recording from it captures what it plays.
            devices = list(pa.get_loopback_device_info_generator())
            try:
                default = pa.get_default_wasapi_loopback()
            except Exception:
                default = None
        else:
            # Microphones: real input devices on the WASAPI host API
            # (Windows' modern audio system), excluding loopback devices.
            wasapi = pa.get_host_api_info_by_type(pyaudio.paWASAPI)
            devices = []
            for i in range(pa.get_device_count()):
                d = pa.get_device_info_by_index(i)
                if (d["hostApi"] == wasapi["index"] and d["maxInputChannels"] > 0
                        and not d.get("isLoopbackDevice", False)):
                    devices.append(d)
            try:
                default = pa.get_device_info_by_index(wasapi["defaultInputDevice"])
            except Exception:
                default = None

        # Find the default device's position in our list (fall back to the first).
        idx = 0 if devices else -1
        if default is not None:
            for i, d in enumerate(devices):
                if d["index"] == default["index"]:
                    idx = i
                    break
        return devices, idx
    finally:
        pa.terminate()


# --------------------------------------------------------------------------- #
#  8. App - the window and the worker threads
# --------------------------------------------------------------------------- #

class App:
    """The application window plus the logic behind its buttons.

    The window is built in _build_ui(). Pressing Start launches a background
    "worker" thread (_live_worker or _file_worker) so the window stays
    responsive while audio is captured and transcribed.
    """

    def __init__(self, root):
        """root - the main Tk window."""
        self.root = root
        root.title(APP_NAME)
        root.geometry("760x660")
        root.minsize(640, 500)

        # Messages from background threads to the GUI thread (see post()).
        self.ui_q = queue.Queue()

        # The loaded Whisper model is kept between sessions, since loading it
        # takes a few seconds. The lock stops two threads loading it at once.
        self.model = None
        self.model_name = None
        self.model_lock = threading.Lock()

        self.stop_event = threading.Event()  # set when the user presses Stop
        self.running = False                 # is a session in progress?
        self.writer = None                   # DocWriter for the current session
        self.last_path = None                # last saved document, for "Open document"
        self.devices = []                    # devices shown in the drop-down
        self.user_chose_path = False         # did the user pick the save path by hand?
        self.words_path = os.path.join(app_dir(), WORDS_FILE_NAME)
        # Settings copied when Start is pressed, for the worker threads (see _start).
        self.session_model = None
        self.session_timestamps = False
        self.session_hotwords = None

        self._build_ui()
        self._on_source_change()            # show the right controls for the default source
        self.root.after(100, self._drain_ui_queue)   # start the message-pump loop
        root.protocol("WM_DELETE_WINDOW", self._on_close)  # confirm before closing mid-session

    # ======================= window layout ======================= #

    def _build_ui(self):
        """Create every widget. The window is a grid with three columns:
        labels on the left, inputs in the (stretchy) middle, buttons on the
        right."""
        pad = {"padx": 8, "pady": 4}
        frm = ttk.Frame(self.root, padding=10)
        frm.pack(fill="both", expand=True)
        frm.columnconfigure(1, weight=1)   # the middle column grows with the window

        # Row 0: "Listen to:" - choice of audio source
        ttk.Label(frm, text="Listen to:").grid(row=0, column=0, sticky="w", **pad)
        self.source = tk.StringVar(value="loopback")
        src = ttk.Frame(frm)
        src.grid(row=0, column=1, columnspan=2, sticky="w")
        for text, val in (("This computer's audio", "loopback"),
                          ("Microphone", "mic"),
                          ("Audio file", "file")):
            ttk.Radiobutton(src, text=text, value=val, variable=self.source,
                            command=self._on_source_change).pack(side="left", padx=(0, 14))

        # Row 1: device drop-down + Refresh (live modes) ...
        self.dev_label = ttk.Label(frm, text="Device:")
        self.dev_label.grid(row=1, column=0, sticky="w", **pad)
        self.device_var = tk.StringVar()
        self.device_combo = ttk.Combobox(frm, textvariable=self.device_var, state="readonly")
        self.device_combo.grid(row=1, column=1, sticky="ew", **pad)
        self.refresh_btn = ttk.Button(frm, text="Refresh", command=self._load_devices)
        self.refresh_btn.grid(row=1, column=2, sticky="ew", **pad)

        # ... or, in the same row, a file path box + Browse (audio-file mode).
        # These are created now but only placed on screen in file mode.
        self.file_var = tk.StringVar()
        self.file_entry = ttk.Entry(frm, textvariable=self.file_var)
        self.browse_btn = ttk.Button(frm, text="Browse…", command=self._browse_audio)

        # Row 2: "Accuracy:" - which Whisper model to use, with a hint
        ttk.Label(frm, text="Accuracy:").grid(row=2, column=0, sticky="w", **pad)
        self.model_var = tk.StringVar(value="small.en")
        model_row = ttk.Frame(frm)
        model_row.grid(row=2, column=1, columnspan=2, sticky="ew")
        self.model_combo = ttk.Combobox(model_row, textvariable=self.model_var,
                                        values=MODELS, state="readonly", width=12)
        self.model_combo.pack(side="left", padx=8, pady=4)
        self.model_note = ttk.Label(model_row, foreground="gray")
        self.model_note.pack(side="left")
        self.model_combo.bind("<<ComboboxSelected>>", lambda e: self._update_model_note())
        self._update_model_note()

        # Row 3: "Save to:" - output .docx path
        ttk.Label(frm, text="Save to:").grid(row=3, column=0, sticky="w", **pad)
        self.out_var = tk.StringVar(value=os.path.join(default_output_folder(), timestamped_name()))
        ttk.Entry(frm, textvariable=self.out_var).grid(row=3, column=1, sticky="ew", **pad)
        ttk.Button(frm, text="Choose…", command=self._choose_output).grid(row=3, column=2, sticky="ew", **pad)

        # Row 4: "Words to expect:" - names and terms to help Whisper spell
        # correctly, separated by commas. Pre-filled with last time's list.
        ttk.Label(frm, text="Words to expect:").grid(row=4, column=0, sticky="w", **pad)
        self.words_var = tk.StringVar(value=self._load_words())
        ttk.Entry(frm, textvariable=self.words_var).grid(row=4, column=1, sticky="ew", **pad)
        ttk.Label(frm, text="names, terms…", foreground="gray").grid(row=4, column=2, sticky="w", **pad)

        # Row 5: timestamps option
        self.ts_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(frm, text="Add timestamps to each paragraph",
                        variable=self.ts_var).grid(row=5, column=1, sticky="w", **pad)

        # Row 6: Start / Stop / Open document
        btns = ttk.Frame(frm)
        btns.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(8, 4))
        self.start_btn = ttk.Button(btns, text="▶  Start listening", command=self._start)
        self.start_btn.pack(side="left", padx=8)
        self.stop_btn = ttk.Button(btns, text="■  Stop", command=self._stop, state="disabled")
        self.stop_btn.pack(side="left", padx=8)
        self.open_btn = ttk.Button(btns, text="Open document", command=self._open_doc, state="disabled")
        self.open_btn.pack(side="right", padx=8)

        # Row 7: live transcript (read-only; we unlock it briefly to add text)
        self.text = scrolledtext.ScrolledText(frm, wrap="word", height=16, font=("Segoe UI", 10))
        self.text.grid(row=7, column=0, columnspan=3, sticky="nsew", padx=8, pady=(6, 4))
        self.text.configure(state="disabled")
        frm.rowconfigure(7, weight=1)   # the transcript box grows with the window

        # Row 8: progress bar (shown only while transcribing a file)
        self.progress = ttk.Progressbar(frm, mode="determinate", maximum=100)

        # Row 9: status line
        self.status_var = tk.StringVar(value="Ready.")
        ttk.Label(frm, textvariable=self.status_var, foreground="#333").grid(
            row=9, column=0, columnspan=3, sticky="w", padx=8, pady=(2, 0))

    def _load_words(self):
        """Read the saved "Words to expect" list, or return '' if there is none."""
        try:
            with open(self.words_path, encoding="utf-8") as f:
                return f.read().strip()
        except OSError:
            return ""

    def _save_words(self, words):
        """Remember the "Words to expect" list for next time. Failing to save
        (e.g. the app is in a read-only folder) is harmless, so it's ignored."""
        try:
            with open(self.words_path, "w", encoding="utf-8") as f:
                f.write(words)
        except OSError:
            pass

    def _hotwords(self):
        """The "Words to expect" text tidied into a single comma-separated
        line for Whisper, or None if the box is empty.

        faster-whisper's `hotwords` option puts these words in front of every
        30-second window Whisper processes, which nudges it towards those
        spellings throughout the whole recording. See also _glossary().
        """
        parts = [w.strip() for w in self.words_var.get().replace("\n", ",").split(",")]
        words = ", ".join(w for w in parts if w)
        return words or None

    def _glossary(self):
        """The same words as a short sentence, 'Glossary: A, B, C.', for
        Whisper's `initial_prompt`, or '' if there are none.

        Testing showed that hotwords alone fix the spelling of names but make
        Whisper leave out full stops, because the hint doesn't look like a
        sentence. Also giving it this properly punctuated line keeps the
        punctuation while still spelling the names correctly.
        """
        return f"Glossary: {self.session_hotwords}." if self.session_hotwords else ""

    def _update_model_note(self):
        """Show the hint for the currently selected model."""
        self.model_note.configure(text=MODEL_NOTES.get(self.model_var.get(), ""))

    def _on_source_change(self):
        """Swap row 1 between the device picker (live modes) and the file
        picker (audio-file mode) when the user changes the source."""
        mode = self.source.get()
        if mode == "file":
            self.device_combo.grid_remove()
            self.refresh_btn.grid_remove()
            self.dev_label.configure(text="Audio file:")
            self.file_entry.grid(row=1, column=1, sticky="ew", padx=8, pady=4)
            self.browse_btn.grid(row=1, column=2, sticky="ew", padx=8, pady=4)
            self.start_btn.configure(text="▶  Transcribe file")
        else:
            self.file_entry.grid_remove()
            self.browse_btn.grid_remove()
            self.dev_label.configure(text="Speakers:" if mode == "loopback" else "Microphone:")
            self.device_combo.grid()
            self.refresh_btn.grid()
            self.start_btn.configure(text="▶  Start listening")
            self._load_devices()

    def _load_devices(self):
        """Fill the device drop-down for the current source (also used by the
        Refresh button, e.g. after plugging in headphones)."""
        mode = self.source.get()
        if pyaudio is None:
            self.device_combo["values"] = ["(live audio needs Windows + PyAudioWPatch)"]
            self.device_combo.current(0)
            self.devices = []
            return
        try:
            self.devices, idx = list_devices(mode)
        except Exception as e:
            self.devices, idx = [], -1
            self._status(f"Couldn't list audio devices: {e}")
        # Hide the technical " [Loopback]" suffix from device names.
        names = [d["name"].replace(" [Loopback]", "") for d in self.devices]
        self.device_combo["values"] = names or ["(no devices found)"]
        self.device_combo.current(idx if idx >= 0 else 0)

    def _browse_audio(self):
        """Let the user pick an audio/video file. Unless they've already
        chosen a save path, name the transcript after the audio file."""
        path = filedialog.askopenfilename(
            title="Choose an audio or video file",
            filetypes=[("Audio/Video", "*.mp3 *.wav *.m4a *.aac *.flac *.ogg *.wma *.mp4 *.mkv *.webm *.mov"),
                       ("All files", "*.*")])
        if path:
            self.file_var.set(path)
            if not self.user_chose_path:
                base = os.path.splitext(os.path.basename(path))[0]
                self.out_var.set(os.path.join(os.path.dirname(self.out_var.get()),
                                              f"{base} - transcript.docx"))

    def _choose_output(self):
        """Let the user pick where to save the Word document."""
        cur = self.out_var.get()
        path = filedialog.asksaveasfilename(
            title="Save transcript as", defaultextension=".docx",
            initialdir=os.path.dirname(cur), initialfile=os.path.basename(cur),
            filetypes=[("Word document", "*.docx")])
        if path:
            self.out_var.set(path)
            self.user_chose_path = True

    # ================= thread-safe GUI updates ================= #

    def post(self, fn, *args):
        """Ask the GUI thread to run fn(*args). Safe to call from any thread;
        this is the ONLY way background threads should touch the window."""
        self.ui_q.put((fn, args))

    def _drain_ui_queue(self):
        """Runs on the GUI thread every 100 ms: execute everything that
        background threads have posted, then schedule itself again."""
        try:
            while True:
                fn, args = self.ui_q.get_nowait()
                fn(*args)
        except queue.Empty:
            pass
        self.root.after(100, self._drain_ui_queue)

    def _status(self, msg):
        """Set the status line at the bottom of the window."""
        self.status_var.set(msg)

    def _append_text(self, text, new_para):
        """Add transcribed text to the on-screen transcript, mirroring the
        paragraph breaks used in the Word document, and scroll to the end."""
        self.text.configure(state="normal")          # unlock to edit
        is_empty = self.text.index("end-1c") == "1.0"
        if new_para and not is_empty:
            self.text.insert("end", "\n\n")
        elif not is_empty:
            self.text.insert("end", " ")
        self.text.insert("end", text)
        self.text.see("end")
        self.text.configure(state="disabled")        # lock again (read-only)

    def _set_progress(self, pct):
        """Show the progress bar at pct percent, or hide it when pct is None."""
        if pct is None:
            self.progress.grid_remove()
        else:
            self.progress.grid(row=8, column=0, columnspan=3, sticky="ew", padx=8, pady=2)
            self.progress["value"] = pct

    # ==================== speech model ==================== #

    def _get_model(self):
        """Return the Whisper model chosen in the Accuracy drop-down when Start
        was pressed, loading it if needed. Called from worker threads.

        The first time a model is used it is downloaded from the internet
        (Hugging Face) into the "models" folder beside the app; afterwards it
        loads from disk and no internet is needed.
        """
        name = self.session_model
        with self.model_lock:
            if self.model is not None and self.model_name == name:
                return self.model   # already loaded
            self.post(self._status, f"Loading the '{name}' speech model… "
                                    "(the first time, it downloads once — this can take a few minutes)")
            # Imported here rather than at the top because it's slow to import,
            # and this way the window appears instantly at start-up.
            from faster_whisper import WhisperModel
            models_dir = os.path.join(app_dir(), "models")
            os.makedirs(models_dir, exist_ok=True)
            # Leave one CPU core free so the window and audio stay responsive.
            threads = max(1, (os.cpu_count() or 4) - 1)
            # "int8" stores the model's numbers as 8-bit integers: much faster
            # on an ordinary CPU, with practically no loss of accuracy.
            self.model = WhisperModel(name, device="cpu", compute_type="int8",
                                      download_root=models_dir, cpu_threads=threads)
            self.model_name = name
            return self.model

    def _language(self):
        """English-only models get language='en'. Multilingual models get
        None, which tells Whisper to detect the language itself."""
        return "en" if self.session_model.endswith(".en") else None

    # ==================== Start / Stop ==================== #

    def _start(self):
        """Start button: check the settings, then launch the right worker
        thread for the chosen source."""
        # --- Check the output path ---
        out = self.out_var.get().strip()
        if not out:
            messagebox.showwarning(APP_NAME, "Choose where to save the Word document.")
            return
        if not out.lower().endswith(".docx"):
            out += ".docx"
            self.out_var.set(out)
        folder = os.path.dirname(out) or "."
        if not os.path.isdir(folder):
            messagebox.showwarning(APP_NAME, f"The folder doesn't exist:\n{folder}")
            return

        # --- Pick the worker for the chosen source ---
        mode = self.source.get()
        if mode == "file":
            path = self.file_var.get().strip()
            if not os.path.isfile(path):
                messagebox.showwarning(APP_NAME, "Choose an audio file to transcribe.")
                return
            target = self._file_worker
            args = (path, out)
        else:
            if pyaudio is None:
                messagebox.showerror(APP_NAME, "Live listening needs Windows with the "
                                               "PyAudioWPatch package installed.")
                return
            i = self.device_combo.current()
            if not self.devices or i < 0 or i >= len(self.devices):
                messagebox.showwarning(APP_NAME, "No audio device selected.")
                return
            target = self._live_worker
            args = (self.devices[i], out)

        # --- Copy the settings for this session ---
        # Worker threads must not read Tkinter widgets or variables (only the
        # GUI thread may), so everything they need is copied here, on the GUI
        # thread, into plain attributes. This also means changing a setting
        # mid-session has no effect until the next Start.
        self.session_model = self.model_var.get()
        self.session_timestamps = self.ts_var.get()
        self.session_hotwords = self._hotwords()
        self._save_words(self.words_var.get().strip())

        # --- Reset the window for a new session ---
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self.text.configure(state="disabled")
        self.stop_event.clear()
        self.running = True
        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.open_btn.configure(state="disabled")
        self.model_combo.configure(state="disabled")   # can't switch model mid-session

        # daemon=True means this thread won't keep the program alive on exit.
        threading.Thread(target=self._guard, args=(target, args), daemon=True).start()

    def _guard(self, target, args):
        """Run a worker, turning any unexpected error into a friendly message,
        and always tidy up the window afterwards."""
        try:
            target(*args)
        except Exception as e:
            traceback.print_exc()   # full details to the console, for debugging
            self.post(messagebox.showerror, APP_NAME, f"Something went wrong:\n\n{e}")
            self.post(self._status, f"Error: {e}")
        finally:
            self.post(self._finished)

    def _stop(self):
        """Stop button: ask the worker to wrap up. It finishes transcribing
        what it already heard before stopping, so nothing is dropped."""
        self.stop_event.set()
        self.stop_btn.configure(state="disabled")
        self._status("Stopping — finishing the last few words…")

    def _finished(self):
        """Runs on the GUI thread when a worker has ended: re-enable the
        controls, make a final save, and report where the file is."""
        self.running = False
        self.start_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")
        self.model_combo.configure(state="readonly")
        self._set_progress(None)

        if self.writer is not None:
            self.open_btn.configure(state="normal")
            if not self.writer.save():
                self._status("Couldn't save — close the document in Word, then press Start again.")
            elif not self.status_var.get().startswith("Error"):
                self._status(f"Done. Saved to {self.writer.path}")
            self.last_path = self.writer.path

        # After a live session, switch to a fresh file name in the same folder
        # so the next session can never overwrite this transcript.
        if self.source.get() != "file":
            self.out_var.set(os.path.join(os.path.dirname(self.out_var.get()), timestamped_name()))
        self.user_chose_path = False

    def _add_text(self, text, new_para, stamp):
        """Called by workers for each piece of transcribed text: write it to
        the document and show it on screen. Warns if Word has the file locked."""
        ok = self.writer.add(text, new_para, stamp)
        self.post(self._append_text, text, new_para)
        if not ok:
            self.post(self._status, "⚠ The document is open in Word, so it can't be saved right now. "
                                    "Close it — the app keeps retrying.")
        return ok

    # ==================== live listening ==================== #

    def _live_worker(self, device, out_path):
        """Background thread for live audio (speakers or microphone).

        Three things run at once during a live session:

          sound device --(raw bytes)--> audio_q --> this thread (Segmenter)
                                                       |
                                              finished utterances
                                                       v
                                       seg_q --> transcriber thread --> Word doc

        Splitting the work this way means listening never pauses while
        Whisper is busy: if transcription falls behind, utterances simply
        wait in seg_q and the status line shows how many are waiting.
        """
        model = self._get_model()
        self.writer = DocWriter(out_path, "Transcript", self.session_timestamps)

        audio_q = queue.Queue()   # raw audio from the sound card
        seg_q = queue.Queue()     # utterances waiting to be transcribed

        def on_segment(audio, new_para, start_s):
            """Segmenter callback: queue a finished utterance."""
            seg_q.put((audio, new_para, start_s))

        def transcriber():
            """Second background thread: turns queued utterances into text."""
            # The last ~200 characters of transcript are given to Whisper as a
            # "prompt" for context (after the glossary of expected words),
            # which helps keep spelling and punctuation consistent across
            # utterances.
            prev_tail = ""
            while True:
                item = seg_q.get()
                if item is None:   # sentinel value: no more audio, finish up
                    break
                audio, new_para, start_s = item
                backlog = seg_q.qsize()
                self.post(self._status, "Listening… (transcribing"
                          + (f", {backlog} waiting)" if backlog else ")"))

                # beam_size=5: consider 5 candidate word sequences (more
                #   accurate than picking one word at a time).
                # vad_filter=False: our Segmenter already removed silence.
                # condition_on_previous_text=False: each utterance is treated
                #   independently, which prevents repeated-text glitches.
                # hotwords / initial_prompt: the user's "Words to expect", so
                #   names are spelled right (see _hotwords and _glossary).
                segments, _ = model.transcribe(
                    audio, language=self._language(), beam_size=5,
                    vad_filter=False, condition_on_previous_text=False,
                    initial_prompt=(self._glossary() + " " + prev_tail).strip() or None,
                    hotwords=self.session_hotwords)

                parts = []
                for s in segments:
                    # Whisper sometimes "hears" words in noise. Skip pieces it
                    # thinks are probably not speech AND has low confidence in.
                    if s.no_speech_prob > 0.6 and s.avg_logprob < -1.0:
                        continue
                    parts.append(s.text.strip())
                text = " ".join(p for p in parts if p).strip()

                if text:
                    self._add_text(text, new_para, fmt_clock(start_s))
                    prev_tail = (prev_tail + " " + text)[-200:]
                if seg_q.empty() and not self.stop_event.is_set():
                    self.post(self._status, "Listening…")

        tthread = threading.Thread(target=transcriber, daemon=True)
        tthread.start()

        # Open the sound device and start segmenting what it hears.
        capture = AudioCapture(device, audio_q)
        seg = Segmenter(capture.rate, on_segment)
        name = device["name"].replace(" [Loopback]", "")
        self.post(self._status, f"Listening to {name}… Saving to {out_path}")
        try:
            while not self.stop_event.is_set():
                try:
                    raw = audio_q.get(timeout=0.25)
                except queue.Empty:
                    # No audio for 0.25 s: the computer is silent (see
                    # Segmenter.feed_silence for why this matters).
                    seg.feed_silence(0.25)
                    continue
                seg.feed(to_mono_float(raw, capture.channels))
        finally:
            # Stop pressed (or an error): release the device, process any
            # audio still queued, send the final utterance, then tell the
            # transcriber to finish and wait for it.
            capture.stop()
            while not audio_q.empty():
                seg.feed(to_mono_float(audio_q.get_nowait(), capture.channels))
            seg.flush()
            seg_q.put(None)
            tthread.join()

    # ==================== audio file ==================== #

    def _file_worker(self, path, out_path):
        """Background thread for transcribing a saved audio/video file.

        Whisper reads and decodes the file itself, so no Segmenter is needed.
        Results arrive one sentence-ish segment at a time, which lets us show
        progress as we go.
        """
        model = self._get_model()
        self.writer = DocWriter(out_path, "Transcript: " + os.path.basename(path),
                                self.session_timestamps)
        self.post(self._status, f"Transcribing {os.path.basename(path)}…")
        self.post(self._set_progress, 0)

        # vad_filter=True: skip long silent stretches in the file (faster, and
        # avoids Whisper inventing words in silence).
        # hotwords / initial_prompt: the user's "Words to expect", so names
        #   are spelled right (see _hotwords and _glossary for why both).
        segments, info = model.transcribe(
            path, language=self._language(), beam_size=5, vad_filter=True,
            vad_parameters={"min_silence_duration_ms": 500},
            hotwords=self.session_hotwords,
            initial_prompt=self._glossary() or None)
        duration = max(info.duration, 0.001)   # avoid dividing by zero

        # `segments` is a generator: the actual transcription happens as we
        # loop over it, so checking stop_event here makes Stop work promptly.
        prev_end = None
        for s in segments:
            if self.stop_event.is_set():
                break
            text = s.text.strip()
            if not text:
                continue
            # New paragraph if there was a long gap since the previous segment.
            new_para = prev_end is None or (s.start - prev_end) > PARAGRAPH_GAP_S
            self._add_text(text, new_para, fmt_clock(s.start))
            prev_end = s.end
            pct = min(100.0, 100.0 * s.end / duration)
            self.post(self._set_progress, pct)
            self.post(self._status, f"Transcribing… {pct:.0f}%  ({fmt_clock(s.end)} of {fmt_clock(duration)})")

    # ==================== misc ==================== #

    def _open_doc(self):
        """Open the last transcript in Word (or the default .docx program)."""
        path = self.last_path
        if path and os.path.exists(path):
            try:
                os.startfile(path)   # Windows-only: open with the default app
            except AttributeError:   # not on Windows: just show where it is
                messagebox.showinfo(APP_NAME, f"Saved to:\n{path}")
        else:
            messagebox.showinfo(APP_NAME, "No document yet.")

    def _on_close(self):
        """Window close button: confirm if a session is running, then stop
        the workers and close."""
        if self.running and not messagebox.askyesno(
                APP_NAME, "Transcription is still running. Stop and quit?\n"
                          "Everything transcribed so far is already saved."):
            return
        self.stop_event.set()
        # Short delay so workers can see the stop signal before the window goes.
        self.root.after(300, self.root.destroy)


# --------------------------------------------------------------------------- #
#  9. Entry point
# --------------------------------------------------------------------------- #

def main():
    """Create the window and run the app until it is closed."""
    if sys.platform == "win32":
        # Tell Windows we handle high-DPI screens ourselves, so text is crisp
        # instead of blurry on scaled displays.
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    root = tk.Tk()
    # Use the native Windows look where available.
    try:
        ttk.Style().theme_use("vista" if sys.platform == "win32" else "clam")
    except tk.TclError:
        pass
    App(root)
    root.mainloop()   # hand control to Tkinter until the window closes


# Only start the app when this file is run directly (not when imported, e.g.
# by tests).
if __name__ == "__main__":
    main()
