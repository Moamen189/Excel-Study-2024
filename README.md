# Excel Schedule Generator

Generate a schedule in Excel with date, work, Science,and tech columns.

## Table of Contents
- [Overview](#overview)
- [Usage](#usage)
- [Installation](#installation)
- [Formatting](#formatting)
- [Contributing](#contributing)

## Overview

This Python script generates a schedule in Excel for a given date range with columns for work, date, and tech. The schedule is formatted with proper alignment, column width, and a table for better visualization.

## Usage

1. Clone the repository:

    ```bash
    git clone https://github.com/yourusername/description-schedule-generator.git
    cd description-schedule-generator
    ```

2. Install the required dependencies:

    ```bash
    pip install -r requirements.txt
    ```

3. Run the script:

    ```bash
    python schedule_generator.py
    ```

4. Check the generated Excel file in the project directory (`2024_Description_Schedule.xlsx`).

## Installation

Ensure you have Python and pip installed. Use the following commands to set up the project:

```bash
git clone https://github.com/yourusername/description-schedule-generator.git
cd description-schedule-generator
pip install -r requirements.txt
```

## Formatting

The script generates an Excel file with the following formatting:

- Date column (Column I): Center-aligned with the format 'yyyy-mm-dd'.
- Work column (Column J): Wrapped text for better readability.
- Tech column (Column L): Wrapped text for better readability.

A table is added to enhance visualization, with alternating row and column stripes.

## Contributing

Contributions are welcome! Follow these steps:

1. Fork the project.
2. Create a new branch: `git checkout -b feature/YourFeature`.
3. Commit your changes: `git commit -am 'Add some feature'`.
4. Push to the branch: `git push origin feature/YourFeature`.
5. Open a pull request.
