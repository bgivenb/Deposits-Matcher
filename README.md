# DepositsMatcher

DepositsMatcher is a standalone application and Python-based tool designed to assist in accounting and bookkeeping tasks. It helps users match deposit amounts between two lists, identifying any discrepancies. 

## Features

- **List Input**: Users can input deposit amounts into two separate lists, List A and List B.
- **Subset Matching**: Finds matching subset totals between the two lists.
- **Discrepancy Detection**: Highlights any unmatched or remaining amounts in both lists.
- **Help Button**: A brief description of the app's functions is accessible via the "Help" button.

## Usage

1. **Input the number of deposits** for each list (List A and List B) in the respective fields.
2. **Click "Generate Deposit Fields"** to create input fields for entering deposit amounts.
3. **Enter deposit values** in each field.
4. **Click "Find Maximum Matching Sum"** to see the result, which shows the maximum matching sum and any discrepancies.

## Installation

### Standalone Executable

A standalone executable version is available, allowing you to run the application without needing Python installed.

### Running from Source

To run from source, you’ll need Python 3 and the following dependencies:
- `tkinter`
- `Pillow`

Install dependencies and run:

```bash
pip install pillow
python depositsmatcher.py
