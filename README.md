# AWS to GCP TCO Mapping Automation Tool 🚀

## 📌 Overview

This project is a Python-based automation tool designed to convert AWS billing PDF files into structured Excel reports while performing TCO (Total Cost of Ownership) mapping from AWS services to their GCP equivalents.

The tool intelligently parses AWS billing PDFs (which may vary in format), extracts service-wise cost and usage data, and maps them to corresponding GCP services for cost comparison and migration analysis.

---

## 🎯 Key Features

* 📄 Automated parsing of AWS billing PDF files
* 📊 Extraction of service-level cost and usage data
* 🔁 Mapping of AWS services to GCP equivalents
* 💰 Cost comparison (AWS vs GCP)
* 📁 Excel report generation (.xlsx format)
* ⚙️ Handles dynamic and varying AWS billing formats

---

## ☁️ AWS Services Covered

The current implementation focuses on the following AWS services:

* Amazon ECR
* Amazon Redshift
* AWS Secrets Manager
* AWS Step Functions
* AWS WAF

> ⚡ The mapping logic can be extended to support additional AWS services.

---

## 🔄 AWS to GCP Mapping Purpose

This tool helps in translating AWS service costs into equivalent GCP services to:

* Estimate migration cost
* Perform cloud cost optimization
* Assist in decision-making for multi-cloud or migration strategies

---

## 📥 Input & 📤 Output

### 📥 Input

* Place AWS billing PDF file inside the `input/` folder
* Example:

```
input/aws_bill.pdf
```

### 📤 Output

* Generates an Excel file in the `output/` folder
* File format: `.xlsx`

### 📊 Output Includes:

* AWS Service Name
* AWS Cost
* Usage Details
* GCP Equivalent Service
* Estimated GCP Cost
* Recommendations (if applicable)

---

## 🏗️ Project Structure

```
.
├── devashish_TCO.py        # Main standalone script
├── requirements.txt        # Python dependencies
├── input/                  # Folder for AWS billing PDFs
├── output/                 # Generated Excel reports
├── README.md               # Project documentation
└── .gitignore              # Ignored files
```

---

## ⚙️ Tech Stack

* **Python**
* **pandas** – Data processing
* **pdfplumber** – PDF parsing
* **openpyxl / xlsxwriter** – Excel generation
* **pytesseract** *(if OCR is used)*

---

## 🚀 Getting Started

### 1️⃣ Clone the Repository

```bash
git clone https://github.com/devashish1711/TCO-mapping-for-any-AWSBilling.pdf-into-.xlsx.git
cd TCO-mapping-for-any-AWSBilling.pdf-into-.xlsx
```

---

### 2️⃣ Install Dependencies

```bash
pip install -r requirements.txt
```

---

### 3️⃣ Add Input File

Place your AWS billing PDF inside the `input/` folder.

---

### 4️⃣ Run the Script

```bash
python devashish_TCO.py
```

---

### 5️⃣ Check Output

* Output Excel file will be generated inside the `output/` folder

---

## 💡 Use Cases

* ☁️ AWS to GCP migration planning
* 💰 Cloud cost optimization
* 📊 Financial analysis of cloud infrastructure
* 📈 TCO comparison reporting

---

## ⚠️ Notes

* AWS billing PDFs may vary in structure — this tool is designed to handle dynamic formats
* Ensure the input PDF contains readable billing data (OCR may be required for scanned PDFs)
* Mapping logic is customizable and can be enhanced based on business needs

---

## 🔮 Future Enhancements

* Support for more AWS services
* Improved GCP cost estimation models
* Web UI / dashboard for visualization
* Integration with AWS Cost Explorer API
* Direct export to Google Sheets

---

## 👨‍💻 Author

**Devashish Prajapat**
Cloud Intern | AWS & GCP Enthusiast

📌 GitHub: https://github.com/devashish1711

---

## ⭐ Support

If you found this project useful, consider giving it a ⭐ on GitHub!

---
