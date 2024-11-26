# Broadcast

Broadcast is a Python package for automating Broadcast Excel Addin operations using xlwings. It simplifies working with financial data, especially when fetching information from Excel add-ins. Please understand that the use of this package assumes you have the Broadcast Excel addin already installed in your pc and that you know the works of the functions in excel, the purpose is just replicate them in python.

## Table of Contents
- [Installation](#installation)
- [Usage](#usage)
- [Contributing](#contributing)
- [Contact](#contact)

## Installation

### Prerequisites
- Python >= 3.7
- xlwings
- pandas

### Install Broadcast
Clone the repository and install the package:
```bash
git clone https://github.com/CorreiaLuan/broadcast.git
cd broadcast
pip install .

or 

pip install git+https://github.com/CorreiaLuan/broadcast.git
```

## Usage

### Example 1: Fetch Data Using the BC Formula
```python
from broadcast import xlAddin

broad = xlAddin()
data = broad.bc(ativo="Ibov", campos=["ult", "drf"])
print(data)
```

### Example 2: Fetch Data Using the BC Formula
```python
from broadcast import xlAddin

broad = xlAddin()
historical_data = broad.bch(
    ativo="Ibov",
    campos=["ult", "drf"],
    data_inicial="2024-01-01",
    data_final="2024-11-26"
)
print(historical_data)
```

## Contributing
Contributions are welcome! Please follow these steps:

1. Fork the repository.
2. Create a new branch for your feature or bugfix:
   ```bash
    git checkout -b feature-name
3. Commit your changes and push the branch:
    git commit -m "Add feature-name"
    git push origin feature-name
4. Submit a pull request to the main branch.

## Contact
Created by [Luan Correia](mailto:luan.a.correialive@gmail.com). Feel free to contact me for questions or feedback.

