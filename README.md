# Project PurestDog

Automation project that downloads videos from Facebook and MagicBrief, converts them into audio and freeze frames, transcribes audio using ChatGPT Whisper-1, and uploads the results to Miro Canvas.

## Tools & Technologies

* Main Tool: UiPath
* Programming Languages: Python, VB.NET
* Python Libraries: pandas, moviepy, requests
* APIs: ChatGPT Whisper-1, Miro API, Google Sheets

## Features

* Automatically fetches video links from Google Sheets and MagicBrief.
* Converts videos into freeze frames for every second using pandas.
* Extracts audio from videos using moviepy.
* Generates transcripts via ChatGPT Whisper-1 API.
* Uploads transcripts and freeze frames to Miro Canvas via API.

## Setup & Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/teebs339/Project-PurestDog.git
   cd Project-PurestDog
   ```
2. Install Python dependencies:

   ```bash
   pip install -r requirements.txt
   ```
3. Configure API keys for Whisper-1 and Miro in the Python scripts.
4. Open Main.xaml in UiPath Studio to view or run the automation workflow.
5. Ensure Excel and Google Sheets used in the workflow are accessible and updated with video links.

## Usage

1. Add video URLs in the Google Sheet or MagicBrief input files.
2. Run the UiPath automation Main.xaml.
3. The bot will:

   * Convert videos into freeze frames
   * Extract audio and generate transcripts
   * Upload outputs to the Miro Canvas
4. Outputs are organized in folders and on Miro for easy access.

## Contributing

Feel free to fork this repo, suggest improvements, or open issues. For questions or feedback, create a GitHub issue.

## License

This project is for educational and demonstration purposes. Make sure you comply with the terms of use for Facebook, MagicBrief, and Miro APIs.
