# X431 to Excel

A modern desktop application built with [Wails](https://wails.io/) for converting diagnostic `.x431` files into professional Excel (`.xlsx`) spreadsheets.

## Key Features

- **Batch Processing**: Drag and drop multiple `.x431` files simultaneously.
- **Smart Start-Time Detection**: Automatically detects the session start time from the filename (e.g., `..._20250709161906_...`). Fallbacks to 1-second increments if no custom end time is set.
- **Dynamic Sampling**: Optional "End Time" picker allows you to set the recording's end time. The app automatically calculates the sampling rate as `(End Time - Start Time) / Total Samples`. (Note: Custom end time applies specifically to the first file in a batch).
- **Locale-Aware Data**: Numeric values are written as native Excel numbers, ensuring correct interpretation regardless of system decimal separators.
- **Robust Parsing**: Hardened backend with panic recovery and input validation to handle malformed or corrupted files gracefully.
- **Detailed Error Reporting**: Specific error messages are displayed directly in the UI if conversion fails.

## Usage

1. **Launch**: Open the application.
2. **Set End Time (Optional)**: If your recording has a specific end time, use the "End Time" picker to set it.
3. **Drop**: Drag your `.x431` files directly into the red drop zone.
4. **Processing**: The app will process each file. If you set an end time, it will be applied to the first file and then cleared automatically.
5. **Done**: Your `.xlsx` files will appear in the source folder.

## Development

This project uses Go for the backend logic and Vanilla HTML/JS/CSS for the frontend.

### Prerequisites

- [Go](https://go.dev/) 1.21+
- [Node.js](https://nodejs.org/) & NPM
- [Wails CLI](https://wails.io/docs/gettingstarted/installation)

### Build Instructions

To build the production executable:

```bash
wails build
```

The resulting binary will be located in `build/bin/`.
