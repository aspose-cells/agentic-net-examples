// Title: C# – Convert HTML to an Encrypted XLSX Workbook with Aspose.Cells
// Description: Loads an HTML file into an Aspose.Cells Workbook, sets a password via workbook.Settings.Password, and saves the workbook as a password‑protected XLSX file. Demonstrates HTML‑to‑Excel conversion and built‑in encryption in .NET.
// Keywords: Aspose.Cells HTML to Excel conversion | C# encrypt Excel workbook | password protect XLSX .NET | LoadOptions Html Aspose.Cells | Workbook.Settings.Password | secure Excel export from HTML
// Common Searches: Aspose.Cells convert HTML to password protected XLSX C# | How to encrypt an Excel file generated from HTML using Aspose.Cells | C# code sample for HTML to encrypted Excel conversion | Set workbook password with Aspose.Cells .NET
// Developer Intent: Generate an Excel file from an HTML source and lock it with a password for secure distribution.
// Use Cases: Create confidential financial reports by converting HTML templates to protected Excel files. | Export web‑page tables to Excel while complying with data‑privacy regulations. | Produce secure invoices or statements from HTML content for client delivery.
// AI Prompts: Write C# code that uses Aspose.Cells to load an HTML file, apply a password, and save as an encrypted XLSX. | Explain which Excel formats support password protection in Aspose.Cells and how workbook.Settings.Password works. | Show how to add error handling for LoadOptions when converting HTML to a protected workbook.

using System;
using Aspose.Cells;

// Loads an HTML file into an Aspose.Cells Workbook, sets a password via workbook.Settings.Password, and saves the workbook as a password‑protected XLSX file. Demonstrates HTML‑to‑Excel conversion and built‑in encryption in .NET.
class HtmlToExcelEncrypt
{
    static void Main()
    {
        // Input HTML file path
        string htmlPath = "input.html";

        // Output Excel file path
        string excelPath = "output.xlsx";

        // Password to protect the workbook
        string password = "Secret123";

        // Load the HTML file into a workbook
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
        Workbook workbook = new Workbook(htmlPath, loadOptions);

        // Apply password protection
        workbook.Settings.Password = password;

        // Save the workbook as an encrypted Excel file
        workbook.Save(excelPath, SaveFormat.Xlsx);
    }
}
