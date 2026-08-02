// Title: C# – Protect an Excel worksheet with a password and allow cell formatting using Aspose.Cells
// Description: Shows how to create a workbook, set AllowFormattingCell to true, protect the first worksheet with a password and all protection types, and save the file as ProtectedWorksheet.xlsx.
// Keywords: Aspose.Cells | C# protect worksheet | Excel password protection | AllowFormattingCell | ProtectionType.All | worksheet protection options | Aspose.Cells example | protect sheet while allowing formatting
// Common Searches: Aspose.Cells protect worksheet password C# | AllowFormattingCell example Aspose.Cells | How to enable cell formatting on a protected sheet using Aspose.Cells | Protect Excel sheet with password but allow formatting Aspose | C# code to protect worksheet and allow formatting
// Developer Intent: Add password protection to a worksheet while keeping cell‑formatting enabled.
// Use Cases: Lock a financial report worksheet with a password but let analysts change fonts, colors, or borders for better readability. | Distribute a template that is read‑only except for header rows where users can adjust styling to match their branding. | Share a data sheet with external partners where the data is secured, yet they can highlight their own notes by formatting cells.
// AI Prompts: Generate C# code using Aspose.Cells to protect a worksheet with a password and enable AllowFormattingCell. | Explain how to protect only selected protection types while still allowing cell formatting in Aspose.Cells for .NET. | Show the steps to unprotect a worksheet, modify the AllowFormattingCell property, and re‑apply protection programmatically.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, set AllowFormattingCell to true, protect the first worksheet with a password and all protection types, and save the file as ProtectedWorksheet.xlsx.
    public class ProtectWorksheetAllowFormattingCell
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Get the protection object for the worksheet
                Protection protection = sheet.Protection;

                // Enable formatting of cells even when the sheet is protected
                protection.AllowFormattingCell = true;

                // Protect the worksheet with a password and all protection types
                sheet.Protect(ProtectionType.All, "myPassword", null);

                // Save the workbook
                workbook.Save("ProtectedWorksheet.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ProtectWorksheetAllowFormattingCell.Run();
        }
    }
}
