// Title: Protect Each Excel Worksheet with Unique Passwords Using Aspose.Cells for .NET
// Description: Loads an XLSX file, derives a base password from the workbook name, then iterates through all worksheets and applies ProtectionType.All with a distinct password (e.g., "Report_1", "Report_2"). The protected workbook is saved as a new file.
// Keywords: Aspose.Cells | C# worksheet protection | Excel password per sheet | generate sheet passwords | ProtectionType.All | file‑name based password | .NET Excel security | Worksheet.Protect overload
// Common Searches: Aspose.Cells protect each sheet with different password | C# set individual passwords for Excel worksheets | generate worksheet passwords from file name Aspose | protect all worksheets programmatically .NET | Worksheet.Protect password example
// Developer Intent: Programmatically assign a separate password to every worksheet in a workbook, using the workbook's filename as the password seed.
// Use Cases: Secure sections of a financial model so that only authorized users can edit specific sheets. | Automate batch processing of multiple reports, ensuring each sheet is locked with a distinct password derived from its source file. | Distribute a template where departmental tabs are protected with passwords like "Template_1", "Template_2" to control access.
// AI Prompts: Generate C# code that loads an Excel file with Aspose.Cells, creates a base password from the file name, protects each worksheet using ProtectionType.All with passwords formatted as '<filename>_<sheetIndex>', and saves the result. | Write a reusable Aspose.Cells method named ProtectSheets(string path) that applies per‑sheet passwords based on the workbook name and returns the path of the protected file.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsWorksheetProtection
{
    // Loads an XLSX file, derives a base password from the workbook name, then iterates through all worksheets and applies ProtectionType.All with a distinct password (e.g., "Report_1", "Report_2"). The protected workbook is saved as a new file.
    class Program
    {
        static void Main()
        {
            // Path to the source workbook
            string sourcePath = "input.xlsx";

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(sourcePath);

            // Derive a base password from the file name (without extension)
            string basePassword = Path.GetFileNameWithoutExtension(sourcePath);

            // Protect each worksheet with a unique password
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];

                // Create a unique password for the worksheet (e.g., "input_1", "input_2", ...)
                string worksheetPassword = $"{basePassword}_{i + 1}";

                // Protect the worksheet with all protection types and the generated password
                // Using the overload Protect(ProtectionType, string, string)
                sheet.Protect(ProtectionType.All, worksheetPassword, null);
            }

            // Save the protected workbook (lifecycle rule: save)
            string outputPath = "output_protected.xlsx";
            workbook.Save(outputPath);
        }
    }
}
