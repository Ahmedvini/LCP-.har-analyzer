
# HAR File to Document Converter

This project provides a tool to convert HAR (HTTP Archive) files into well-structured Word documents. It extracts request and response details from the HAR file and generates a formatted report that includes URL, HTTP method, response status, response time, and response size.

### Features

* Parse **HAR** files to extract HTTP request and response details.
* Generate a **Word document (.docx)** with request and response data.
* User-friendly **GUI** using Tkinter to select the HAR file and output destination.
* Performance analysis of LCP (Largest Contentful Paint) time and its classification.

### Requirements

To run this project, you need to have the following libraries installed:

* **Python 3.x**
* **pandas** (for potential data manipulation or performance tracking)
* **python-docx** (for generating Word documents)
* **Tkinter** (for the graphical user interface)

You can install the required libraries using:

```bash
pip install pandas python-docx
```

### How to Run

1. Clone the repository:

   ```bash
   git clone https://github.com/Ahmedvini/LCP-.har-analyzer.git
   cd LCP-.har-analyzer

   ```

2. Install the required dependencies:

   ```bash
   pip install pandas python-docx
   ```

3. Run the script:

   ```bash
   python har_analyzer.py
   ```

4. Use the GUI to select your HAR file and output location for the generated document.

### How It Works

1. **Parsing the HAR File**:
   The tool opens and parses the `.har` file, extracting key details such as the URL, HTTP method, response status, response time, and response size from each entry in the HAR log.

2. **Generating the Word Document**:
   The parsed data is then used to generate a Word document (.docx) with a formatted table, allowing users to view request-response details in an easy-to-read manner.

3. **Performance Analysis**:
   The tool checks the **Largest Contentful Paint (LCP)** time from the HAR data and classifies the performance based on Google’s thresholds:

   * **Good**: LCP ≤ 2.5 seconds
   * **Needs Improvement**: 2.5s < LCP ≤ 4.0s
   * **Poor**: LCP > 4.0s
     This performance classification is added to the generated document for further insights.

### Example of Output

The output Word document will include:

* An **overview** of the analysis.
* A **table** with the following columns:

  * **URL**
  * **Method**
  * **Status**
  * **Response Time (ms)**
  * **Response Size (bytes)**
* **LCP Performance Classification** based on the largest contentful paint data from the HAR file.

### Troubleshooting

1. **Missing dependencies**: Make sure you’ve installed the required libraries using the `pip install` command.

2. **Empty HAR File**: If the HAR file is empty or contains errors, the script will notify you with an error message.

3. **Word Document Formatting**: If the generated Word document looks misaligned or corrupt, try opening it with a different version of Microsoft Word or LibreOffice.

### Contributing

Contributions are welcome! If you have suggestions, improvements, or bug fixes, feel free to fork the repository and submit a pull request.

1. Fork the repository.
2. Create a new branch (`git checkout -b feature-xyz`).
3. Commit your changes (`git commit -am 'Add new feature'`).
4. Push to the branch (`git push origin feature-xyz`).
5. Create a new pull request.

### License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

### Acknowledgements

* **pandas**: For handling any data manipulation.
* **python-docx**: For generating Word documents.
* **Tkinter**: For creating the graphical user interface.

---
