# Office365 SharePoint PowerPoint Downloader

A Python script to automatically capture PowerPoint presentations from SharePoint sites and convert them to PDF. This project does not require SharePoint editing permissions or logging into Microsoft Office 365.

## Easy Start

- For an easy start, click [here](easy-start-guide.md).

## 中文版

- [点击这里查看中文版 README](README-zh.md)

## Features

- Automatically navigates to SharePoint PowerPoint presentations
- Captures each slide as a high-quality screenshot
- Combines all slides into a single PDF file
- Handles animations and transitions
- Supports full-screen presentation mode
- Intelligent last slide detection

## Prerequisites

- Python 3.7 or higher
- Chrome browser
- Required Python packages:
  ```
  selenium
  Pillow
  ```
  This project use  
  ```bash
  pip install -r requirements.txt
  ```

## Installation

1. Clone this repository or download the script
2. Install the required packages:
   ```bash
   pip install selenium Pillow
   ```
3. Make sure you have Chrome browser installed

## Usage

1. Run the script:
   ```bash
   python powerpoint_capture-en.py
   ```

2. When prompted, enter:
   - The SharePoint PowerPoint URL
   - Output folder name (optional, defaults to 'slides')

3. The script will:
   - Open the presentation in Chrome
   - Enter presentation mode
   - Capture each slide
   - Save all slides as a PDF

## Output

- Individual slide screenshots are saved in the specified folder (default: 'slides')
- A combined PDF file is created in the same folder
- Debug information is saved if any errors occur

## Error Handling

The script includes robust error handling:
- Saves page source when errors occur
- Retries failed operations
- Provides detailed logging
- Handles network issues and page loading delays

## Known Limitations

- Requires active Office365 login session in Chrome
- May not capture some complex animations
- Network connection required
- Screen resolution should be at least 1920x1080 for best results

## Troubleshooting

If you encounter issues:
1. Check your Office365 login status
2. Ensure stable internet connection
3. Check the error logs and page source in the output folder
4. Verify you have the correct permissions to access the PowerPoint

## License

This project is licensed under the MIT License - see the LICENSE file for details.
