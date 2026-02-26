# X431 to Excel

A modern desktop application built with [Wails](https://wails.io/) for converting diagnostic `.x431` files into professional Excel (`.xlsx`) spreadsheets.

## Key Features

- **Batch Processing**: Drag and drop multiple `.x431` files simultaneously for rapid conversion.
- **Smart Timestamps**: Automatically detects the session start time from the filename (e.g., `..._20250709161906_...`) and generates second-by-second timestamps for every data point.
- **Locale-Aware Data**: Numeric values are written as native Excel numbers, ensuring correct interpretation regardless of whether your system uses periods or commas as decimal separators.
- **Automatic Organization**: Resulting Excel files are saved in the same directory as the source files, with each sheet named after the capture time.
- **Modern UI**: A clean, responsive, and high-performance interface with a premium red-themed design.

## Usage

1. **Launch**: Open the application.
2. **Drop**: Drag your `.x431` files from your file explorer directly into the red drop zone.
3. **Wait**: The app will process each file sequentially.
4. **Done**: Your `.xlsx` files will appear in the source folder ready for analysis.

## Development

This project uses Go for the backend logic and Vanilla HTML/JS/CSS for the frontend, bundled via Wails.

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
