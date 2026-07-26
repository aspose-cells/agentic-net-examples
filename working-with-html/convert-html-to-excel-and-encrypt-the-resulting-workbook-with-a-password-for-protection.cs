// Title: Convert HTML to a Password‑Protected Excel Workbook with Aspose.Cells for .NET
// Description: This C# example shows how to load an HTML file into an Aspose.Cells Workbook, apply workbook‑level password protection, and save the result as an encrypted XLSX document.
// Keywords: Aspose.Cells | HTML to Excel conversion | C# password protection | Workbook.Settings.Password | encrypt XLSX | secure Excel export | convert HTML to XLSX
// Common Searches: Aspose.Cells convert HTML to XLSX C# | how to password protect Excel file using Aspose.Cells | C# load HTML and save encrypted workbook | set workbook password Aspose.Cells .NET | encrypt Excel generated from HTML
// Developer Intent: Create an Excel file from an HTML source and lock it with a password.
// Use Cases: Export web‑based invoices to a protected XLSX file for confidential distribution. | Automate regulatory‑compliant reporting by converting HTML tables into encrypted workbooks. | Provide downloadable Excel sheets from a portal where each file is secured with a known password.
// AI Prompts: Generate C# code that uses Aspose.Cells to read an HTML file, assign a workbook password, and write an encrypted XLSX file. | Explain the limitations of Workbook.Settings.Password for protecting Excel files in Aspose.Cells. | Show how to combine workbook password protection with worksheet‑level protection after converting HTML to Excel.

using System;
using Aspose.Cells;

namespace HtmlToExcelEncryptionDemo
{
    // This C# example shows how to load an HTML file into an Aspose.Cells Workbook, apply workbook‑level password protection, and save the result as an encrypted XLSX document.
    class Program
    {
        static void Main()
        {
            // Path to the source HTML file
            string htmlPath = "input.html";

            // Path for the encrypted Excel file to be saved
            string excelPath = "output.xlsx";

            // Password to protect the workbook
            string password = "SecurePassword123";

            // Load the HTML file into a Workbook object
            // Aspose.Cells automatically parses the HTML and creates worksheets
            Workbook workbook = new Workbook(htmlPath);

            // Apply password protection using the WorkbookSettings.Password property (rule)
            workbook.Settings.Password = password;

            // Save the workbook as an Excel file (XLSX format)
            workbook.Save(excelPath, SaveFormat.Xlsx);

            Console.WriteLine($"HTML file '{htmlPath}' has been converted to Excel and saved as '{excelPath}' with password protection.");
        }
    }
}
