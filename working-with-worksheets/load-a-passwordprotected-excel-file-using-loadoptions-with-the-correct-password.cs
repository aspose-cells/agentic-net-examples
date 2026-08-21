// Title: Load a password‑protected Excel workbook with Aspose.Cells for .NET and save it unprotected
// Description: Shows how to open an encrypted .xlsx file using Aspose.Cells LoadOptions.Password, read a cell to confirm access, remove the opening password, and write a new workbook without protection. Includes file‑existence validation and exception handling.
// Keywords: Aspose.Cells load password protected workbook | LoadOptions.Password .NET | open encrypted Excel file C# | remove workbook password Aspose.Cells | save unprotected Excel Aspose | C# Aspose.Cells example | Excel file security Aspose
// Common Searches: How to open a password protected Excel file with Aspose.Cells | Aspose.Cells LoadOptions password example | Remove opening password from Excel using Aspose | Save unprotected copy of protected workbook C# | Aspose.Cells read cell from encrypted workbook
// Developer Intent: Open a secured Excel file, access its data, and optionally create an unprotected version using Aspose.Cells for .NET.
// Use Cases: Extract data from a protected report before analysis. | Automate batch de‑cryption of multiple workbooks for downstream processing. | Validate workbook contents before applying updates or adding worksheets.
// AI Prompts: Generate C# code that opens a password‑protected .xlsx file with Aspose.Cells, reads a specific cell, and handles incorrect passwords gracefully. | Provide a snippet to clear the opening password of a loaded workbook and save it as a new unprotected file using Aspose.Cells. | Explain the steps to verify a protected workbook’s content, remove its password, and export an unencrypted copy in a .NET application.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to open an encrypted .xlsx file using Aspose.Cells LoadOptions.Password, read a cell to confirm access, remove the opening password, and write a new workbook without protection. Includes file‑existence validation and exception handling.
    public class LoadPasswordProtectedWorkbook
    {
        public static void Run()
        {
            // Path to the password‑protected Excel file
            string filePath = "protected.xlsx";

            // Verify that the source file exists
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: File \"{filePath}\" not found.");
                return;
            }

            try
            {
                // Create LoadOptions and set the password required to open the workbook
                LoadOptions loadOptions = new LoadOptions
                {
                    Password = "test"
                };

                // Load the workbook using the LoadOptions
                Workbook workbook = new Workbook(filePath, loadOptions);

                // Verify that the workbook was loaded (e.g., read a cell value)
                Console.WriteLine("Cell A1 value: " + workbook.Worksheets[0].Cells["A1"].Value);

                // Optional: remove the password protection and save an unprotected copy
                workbook.Settings.Password = null; // clear the opening password
                string unprotectedPath = "unprotected.xlsx";
                workbook.Save(unprotectedPath);
                Console.WriteLine($"Unprotected workbook saved as \"{unprotectedPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            LoadPasswordProtectedWorkbook.Run();
        }
    }
}
