# GitHub Pages XLSX Attendance

This project is a static web application that allows users to upload an XLSX file, process the data to remove unnecessary columns, and calculate attendance ratios. The processed data can then be downloaded.

## Project Structure

```
github-pages-xlsx-attendance
├── index.html          # Main HTML page with the upload form
├── src
│   ├── app.js         # Main JavaScript file for handling file uploads and processing
│   ├── utils
│   │   └── xlsxProcessor.js  # Utility functions for processing XLSX files
│   └── styles
│       └── main.css   # CSS styles for the application
└── README.md          # Documentation for the project
```

## Features

- Upload XLSX files
- Remove columns with all zeros
- Retain the "Presenze" column (kept as a COUNTIF formula)
- Calculate and add a "Percentuale" column for attendance ratios
- Download the processed file

## Usage (locally)

1. Clone the repository to your local machine.
2. Serve the folder locally by issuing the following command
    ```python3 -m  http.server 8000```
3. Then open your browser at `http://localhost:8000`
4. Use the form to upload your XLSX file (e.g., `Presenze_corso.xlsx`).
5. The app will remove date columns that are entirely zero, preserve the "Presenze" formula to count 1's over remaining dates, and add a "Percentuale" column.
6. A processed `.xlsx` file will be downloaded automatically.

## Usage (online)
1. Head over to [introcp.github.io/backoffice-fixed/](introcp.github.io/backoffice-fixed/)
2. Upload your XLSX file using the provided form
3. The app will process the file as described above and download the modified file.

## Dependencies

This project uses the following libraries:

- [SheetJS (xlsx)](https://github.com/SheetJS/sheetjs) - For reading and processing XLSX files.

Make sure to include the necessary scripts in your `index.html` to utilize these libraries.

<!-- ## License

This project is licensed under the MIT License. -->

<!-- ## How to run locally
```python
python3 -m  http.server 8000
```
Then open your browser at `http://localhost:8000` -->

<!-- ## Deploying to GitHub Pages

- Commit and push to the repository.
- In GitHub, go to Settings → Pages.
- Under "Build and deployment", select the `main` (or default) branch and `/ (root)` folder.
- Save. After a few minutes, the site will be available at the provided Pages URL. -->