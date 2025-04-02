import os
import sys
import subprocess
import site
import getpass
from pathlib import Path
import win32com.client
import setuptools
from setuptools import setup

# Define project metadata and install Python dependencies
setup(
    name="VoiceVoyager",
    version="1.0.0",
    author="skapezMpier",
    author_email="skapezmpier@example.com",
    description="A comprehensive audio transcription and analysis tool",
    long_description=open("README.md").read(),
    long_description_content_type="text/markdown",
    url="https://github.com/skapezMpier/voicevoyager",
    py_modules=["voicevoyager"],
    install_requires=[
        "pydub",
        "ttkbootstrap",
        "google-generativeai",
        "speechrecognition",
        "openai",
        "python-docx",
        "reportlab",
        "pygame",
        "cryptography",
        "openai-whisper;platform_system=='Windows'",
    ],
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: Custom License",
        "Operating System :: Microsoft :: Windows",
    ],
)

def run_command(command, shell=True):
    """Run a shell command and handle errors."""
    try:
        subprocess.run(command, shell=shell, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        return True
    except subprocess.CalledProcessError as e:
        print(f"Error running command {command}: {e}")
        return False

def install_choco():
    """Install Chocolatey if not already installed."""
    print("Checking for Chocolatey...")
    if run_command("choco -v"):
        print("Chocolatey is already installed.")
        return True
    print("Installing Chocolatey...")
    install_cmd = (
        'powershell -NoProfile -ExecutionPolicy Bypass -Command "'
        'iex ((New-Object System.Net.WebClient).DownloadString(\'https://chocolatey.org/install.ps1\'))"'
    )
    if run_command(install_cmd):
        print("Chocolatey installed successfully.")
        # Add Chocolatey to PATH
        os.environ["PATH"] += os.pathsep + r"C:\ProgramData\chocolatey\bin"
        return True
    else:
        print("Failed to install Chocolatey. Please install it manually: https://chocolatey.org/install")
        return False

def install_ffmpeg():
    """Install FFmpeg using Chocolatey if not already installed."""
    print("Checking for FFmpeg...")
    if run_command("ffmpeg -version"):
        print("FFmpeg is already installed.")
        return True
    if not install_choco():
        return False
    print("Installing FFmpeg via Chocolatey...")
    if run_command("choco install ffmpeg -y"):
        print("FFmpeg installed successfully.")
        # Add FFmpeg to PATH
        os.environ["PATH"] += os.pathsep + r"C:\ProgramData\chocolatey\lib\ffmpeg\tools"
        return True
    else:
        print("Failed to install FFmpeg. Please install it manually: choco install ffmpeg")
        return False

def create_desktop_shortcut():
    """Create a desktop shortcut for VoiceVoyager with the app icon."""
    print("Creating desktop shortcut...")
    try:
        # Get the desktop path
        user = getpass.getuser()
        desktop = Path(f"C:/Users/{user}/Desktop")
        shortcut_path = desktop / "VoiceVoyager.lnk"

        # Get the path to voicevoyager.py
        script_path = Path(__file__).parent / "voicevoyager.py"
        script_path = script_path.resolve()

        # Get the path to the Python executable
        python_exe = sys.executable

        # Get the path to the icon
        icon_path = Path(__file__).parent / "icon.ico"
        icon_path = icon_path.resolve()

        # Create the shortcut
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(str(shortcut_path))
        shortcut.TargetPath = python_exe
        shortcut.Arguments = f'"{script_path}"'
        shortcut.WorkingDirectory = str(script_path.parent)
        shortcut.IconLocation = str(icon_path)
        shortcut.Description = "VoiceVoyager - Audio Transcription and Analysis Tool"
        shortcut.save()
        print(f"Desktop shortcut created at {shortcut_path}")
    except Exception as e:
        print(f"Failed to create desktop shortcut: {e}")

def main():
    print("Setting up VoiceVoyager...")

    # Install FFmpeg
    if not install_ffmpeg():
        print("Setup cannot proceed without FFmpeg. Exiting.")
        sys.exit(1)

    # Create desktop shortcut
    create_desktop_shortcut()

    print("Setup completed successfully! You can now launch VoiceVoyager from the desktop shortcut.")

if __name__ == "__main__":
    main()