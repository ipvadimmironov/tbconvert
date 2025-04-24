# Excel File Merger

A simple Node.js application that merges multiple Excel (.xlsx) files into a single Excel file.

## Features

- Reads all Excel files from the `Source` directory
- Merges data from all sheets in each file
- Creates a new sheet for each original sheet with a unique name
- Saves the merged data to a new Excel file

## Requirements

- Node.js
- npm

## Installation

```bash
npm install
```

## Usage

1. Place your Excel (.xlsx) files in the `Source` directory
2. Run the application:

```bash
node index.js
```

3. Find the merged file at `merged_data.xlsx` in the project root directory 