# VoiceVoyager

![VoiceVoyager Logo](icon.png) <!-- Replace with your actual logo/icon if available -->

**VoiceVoyager** is a powerful audio transcription and analysis tool designed to transform your audio files into actionable insights with ease. Whether you're transcribing meetings, lectures, interviews, or podcasts, VoiceVoyager provides a user-friendly interface to transcribe, analyze, and export your audio content efficiently. Built with flexibility in mind, it supports multiple transcription models, advanced analysis features, and interactive playback controls.

---

## Features

- **Multiple Transcription Models**:
  - **Gemini**: Uses Whisper for initial transcription, then refines with Gemini for speaker labels and clarity (online).
  - **Google Speech**: Leverages Google Speech-to-Text API for transcription (online).
  - **Whisper**: Uses OpenAI’s Whisper model for high-quality transcription (online).
  - **PocketSphinx (Offline)**: Offline transcription for English audio.
  - **Whisper Local (Offline)**: Offline transcription using the `openai-whisper` library (requires installation).

- **Advanced Analysis Tools** (Powered by Gemini):
  - Translate transcriptions to other languages.
  - Extract key phrases and words.
  - Detect action items or tasks.
  - Tag non-speech events (e.g., laughter, applause).
  - Ask questions about the transcription.
  - Perform sentiment analysis with emoji-enhanced output.

- **Interactive Audio Playback**:
  - Play, stop, and seek through audio with a slider.
  - Displays current position and selected duration.

- **Export Options**:
  - Export transcriptions as professional PDF or DOCX files.

- **Secure API Key Management**:
  - API keys for Gemini and OpenAI are encrypted and stored securely.

- **Customizable Settings**:
  - Choose transcription chunk size (5-300 seconds).
  - Normalize audio for better transcription accuracy.
  - Switch between light and dark themes.

- **Cross-Platform Support**:
  - Currently optimized for Windows, with plans for macOS and Linux support.

---

## Requirements

- **Operating System**: Windows 10 or later (macOS and Linux support coming soon)
- **Python**: 3.8 or higher
- **FFmpeg**: Required for audio processing (installation instructions provided below)
- **Internet Connection**: Required for online transcription models (Gemini, Google Speech, Whisper)
- **API Keys**:
  - Gemini API key (for Gemini transcription and analysis)
  - OpenAI API key (for Whisper transcription)
  - Configure API keys in the "API" tab of the app

---

## Installation (Windows)

### Prerequisites
- **Python 3.8 or Higher**: Download and install from [python.org](https://www.python.org/downloads/). Ensure `pip` is installed and added to your PATH.
- **Git**: Optional, for cloning the repository. Download from [git-scm.com](https://git-scm.com/downloads).

### Steps
1. **Clone or Download the Repository**:
   - Clone the repository using Git:
     ```bash
     git clone https://github.com/skapezMpier/voicevoyager.git
     cd voicevoyager
     ```
   - Alternatively, download the ZIP file from GitHub and extract it.

2. **Install FFmpeg**:
   - FFmpeg is required for audio processing. Install it using Chocolatey (a Windows package manager):
     - Open a Command Prompt or PowerShell as Administrator.
     - Install Chocolatey by running:
       ```powershell
       Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
       ```
     - Install FFmpeg:
       ```bash
       choco install ffmpeg -y
       ```
   - If you don’t want to use Chocolatey, download FFmpeg manually from [ffmpeg.org](https://ffmpeg.org/download.html) and add it to your system PATH.

3. **Install Python Dependencies**:
   - Navigate to the project directory:
     ```bash
     cd voicevoyager
     ```
   - Install the required packages:
     ```bash
     pip install -r requirements.txt
     ```
   - If `requirements.txt` is not present, install the following packages manually:
     ```bash
     pip install pydub ttkbootstrap google-generativeai speechrecognition openai python-docx reportlab pygame cryptography
     ```
   - For offline Whisper transcription, install the `openai-whisper` package:
     ```bash
     pip install openai-whisper
     ```

4. **Launch VoiceVoyager**:
   - Run the application:
     ```bash
     python voicevoyager.py
     ```
   - The VoiceVoyager window should open, ready for use.

---

## Usage

1. **Select an Audio File**:
   - In the "Transcription" tab, click "Browse" to select an audio file (.mp3, .wav, .aiff, .flac).

2. **Configure Settings**:
   - **Model**: Choose a transcription model (e.g., Gemini, Google Speech, Whisper, PocketSphinx, or Whisper Local).
   - **Language**: Select the audio language (e.g., en-US, es-ES, fr-FR).
   - **Chunk Size**: Adjust the duration of each transcription segment (5-300 seconds).
   - **Normalize Audio**: Enable to adjust audio volume for better transcription accuracy.

3. **Transcribe**:
   - Click "Transcribe" to start the transcription process.
   - The transcription will appear in the output window, with timestamps for each segment.
   - A text file with the transcription will be saved in the same directory as the audio file.

4. **Playback Audio**:
   - Use the "Play", "Stop", and seek slider to listen to the audio.
   - The current position and selected duration are displayed in real-time.

5. **Analyze**:
   - Switch to the "Analysis" tab to perform advanced analysis:
     - **Translate**: Convert the transcription to another language.
     - **Extract Keywords**: Identify key phrases or words.
     - **Detect Actions**: Find action items or tasks.
     - **Tag Events**: Mark non-speech events (e.g., laughter, applause).
     - **Ask Question**: Query the transcription for specific information.
     - **Sentiment Analysis**: Analyze the tone with emoji-enhanced output.

6. **Export**:
   - Click "Export" in the "Analysis" tab to save the transcription as a PDF or DOCX file.

7. **Configure API Keys**:
   - In the "API" tab, enter your Gemini and OpenAI API keys.
   - Click "Save API Keys" to store them securely (keys are encrypted).

---

## Shortcuts
- **Ctrl+O**: Open an audio file
- **Ctrl+Q**: Exit the application

---

## Dependencies
VoiceVoyager relies on the following Python packages:
- `pydub`: Audio processing
- `ttkbootstrap`: Modern GUI toolkit
- `google-generativeai`: Gemini API for transcription and analysis
- `speechrecognition`: Google Speech and PocketSphinx transcription
- `openai`: Whisper transcription
- `python-docx`: DOCX export
- `reportlab`: PDF export
- `pygame`: Audio playback
- `cryptography`: Secure API key storage
- `openai-whisper`: Optional, for offline Whisper transcription

Additionally, FFmpeg is required for audio processing.

---

## Troubleshooting

- **FFmpeg Not Found**:
  - Ensure FFmpeg is installed and added to your system PATH.
  - Run `ffmpeg -version` in a command prompt to verify.
  - If missing, install it using Chocolatey (`choco install ffmpeg -y`) or manually.

- **No Internet Connection**:
  - Online models (Gemini, Google Speech, Whisper) require an internet connection.
  - Use offline models (PocketSphinx, Whisper Local) if you’re offline.

- **API Key Errors**:
  - Ensure valid Gemini and OpenAI API keys are entered in the "API" tab.
  - If you don’t have API keys, use offline models (PocketSphinx, Whisper Local).

- **Whisper Local Not Working**:
  - Ensure the `openai-whisper` package is installed (`pip install openai-whisper`).
  - The first run may download the Whisper model, which requires an internet connection.

- **Audio Playback Issues**:
  - Ensure FFmpeg is installed correctly.
  - Try a different audio file format (.wav is recommended for compatibility).

---

## License

VoiceVoyager is released under the **AreSistv(Assistive) Technology Software License** (Non-Commercial Open Source License). Key points:
- Free to use, modify, and distribute for personal or research purposes.
- Can be integrated into commercial products, but must remain freely accessible.
- Attribution to the original author (skapezMpier) is required.

See the [LICENSE](LICENSE) file for full details.

---

## Contributing

Contributions are welcome! To contribute:
1. Fork the repository.
2. Create a new branch for your feature or bug fix.
3. Submit a pull request with a detailed description of your changes.

Please ensure your contributions align with the license terms.

---

## Contact

For support, feedback, or inquiries, reach out to **skapezMpier**:
- Email: [skapezempire@gmail.com](mailto:skapezempire@gmail.com)
- LinkedIn: [skapezMpier](https://www.linkedin.com/company/skapezmpier/)

---

## Roadmap

- Add support for macOS and Linux.
- Improve audio playback with pause/resume functionality.
- Add batch processing for multiple audio files.
- Enhance offline transcription capabilities.
- Create a standalone executable using PyInstaller.

---

**Created by skapezMpier, March 2025**


