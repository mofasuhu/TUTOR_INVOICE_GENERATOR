# Tutor Invoice Generator

The Tutor Invoice Generator is a Python application designed to process tutor data from an Excel file (`tutorlist.xlsx`) and generate consolidated data for each tutor. The application then creates individual PDF invoices for each tutor based on their teaching sessions, special tasks, and other compensation details.

## Project Structure

```plaintext
TUTOR_INVOICE_GENERATOR/
│
├─── fonts/
│    ├── NotoNaskhArabic-Bold.ttf
│    ├── NotoNaskhArabic-Regular.ttf
│    └── NotoSerif-Bold.ttf
│
├─── PDFs/  # This folder will contain the generated PDF invoices
│
├─── SampleOutputsToCheck/
│    ├── consolidated_tutor_data.xlsx
│    └── PDFs/  # This folder contains the sample generated PDF invoices that I prepared for you
│
├── invoices_app.py
├── repair.bat
├── requirements.txt
├── RUN.bat
├── setup.bat
├── .gitignore
├── README.md
├── LICENSE
└── tutorlist.xlsx
```

## Features

- **Data Processing:**
  - Reads tutor data from `tutorlist.xlsx`.
  - Consolidates data for each tutor, including sessions, prices, special tasks, CRM payments, and compensation.
  - Generates a new Excel file (`consolidated_tutor_data.xlsx`) with the consolidated data for all tutors.

- **Invoice Generation:**
  - Creates a PDF invoice for each tutor using the consolidated data.
  - Uses ReportLab to format the PDF invoices, including sections for tutor information, payment details, and a detailed table of compensations.

- **Handling Arabic Text:**
  - Utilizes `arabic-reshaper` and `python-bidi` to correctly display Arabic text in the generated PDF invoices.

## Prerequisites

- Python 3.x
- Required Python packages listed in `requirements.txt`

## Installation

1. Clone the repository:
    ```
    git clone https://github.com/mofasuhu/TUTOR_INVOICE_GENERATOR.git
    ```

2. Navigate to the project directory:
    ```
    cd TUTOR_INVOICE_GENERATOR
    ```

3. Run the `setup.bat` file to create a virtual environment and install dependencies:
    ```
    setup.bat
    ```

4. Add any new modules to the \`requirements.txt\` file and run:
    ```
    repair.bat
    ```

## Usage

1. To run the main script and generate invoices:
    ```bash
    RUN.bat
    ```
    
2. The application will read data from `tutorlist.xlsx`, process it, and generate consolidated data in `consolidated_tutor_data.xlsx`. It will also create individual PDF invoices for each tutor in the `PDFs` directory.

## Requirements

The required Python packages are listed in \`requirements.txt\`:

```plaintext
reportlab==4.2.0
pandas==2.1.1
fpdf==1.7.2
python-bidi==0.4.2
arabic-reshaper==3.0.0
openpyxl==3.1.2
```

## Files

- **invoices_app.py:** Main script to process data and generate invoices.
- **requirements.txt:** Lists all Python dependencies needed for the project.
- **setup.bat:** Sets up the virtual environment and installs dependencies.
- **RUN.bat:** Runs the \`invoices_app.py\` script.
- **repair.bat:** Re-installs dependencies in case of updates to \`requirements.txt\`.
- **tutorlist.xlsx:** Source Excel file containing raw tutor data.
- **consolidated_tutor_data.xlsx:** Output Excel file with consolidated data (sample provided in \`SampleOutputsToCheck/\`).
- **fonts/**: Directory containing necessary font files for Arabic text.
- **PDFs/**: Directory where generated PDF invoices are saved.
- **.gitignore:** Specifies files and directories to be ignored by Git.
- **README.md:** This file, providing an overview and instructions for the project.
- **LICENSE:** The license under which this project is distributed.

## License

This project is licensed under the MIT License.

## Contributing

1. Fork the repository.
2. Create a new branch.
3. Make your changes.
4. Submit a pull request.

## Contact

If you have any questions or suggestions, feel free to contact us at [mofasuhu@gmail.com](mailto:mofasuhu@gmail.com).

---

This project was created to simplify the process of generating invoices for tutors based on their sessions and compensation details.
