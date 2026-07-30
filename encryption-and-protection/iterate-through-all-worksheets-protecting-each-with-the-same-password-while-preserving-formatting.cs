// Title: C# – Protect Every Worksheet in Aspose.Cells with a Single Password (Formatting Preserved)
// Description: Demonstrates how to iterate through all worksheets in a workbook and apply the same password using Worksheet.Protect(ProtectionType.All) while leaving cell styles and formatting untouched, then saves the file as ProtectedWorksheets.xlsx.
// Keywords: Aspose.Cells worksheet protection C# | protect all sheets Aspose.Cells | Worksheet.Protect password | preserve formatting Excel protection | bulk worksheet lock .NET
// Common Searches: protect all worksheets with one password Aspose.Cells | keep cell formatting when protecting Excel sheets C# | apply same password to multiple worksheets Aspose | Worksheet.Protect all sheets example | bulk protect Excel workbook Aspose.Cells
// Developer Intent: Apply identical password protection to every worksheet in a workbook without altering any formatting.
// Use Cases: Generate a multi‑sheet report and lock all sheets with a shared password before distribution. | Load an existing Excel file, enforce read‑only protection on each sheet, and retain original styles. | Create a template workbook where all worksheets are pre‑protected to prevent accidental edits.
// AI Prompts: Show C# code that protects every worksheet in an Aspose.Cells workbook with one password while keeping formatting. | How can I load an existing Excel file and apply the same protection to all its sheets using Aspose.Cells for .NET? | Explain the parameters of Worksheet.Protect and how to use ProtectionType.All to preserve cell styles.

using System;
using Aspose.Cells;

namespace WorksheetProtectionDemo
{
    // Demonstrates how to iterate through all worksheets in a workbook and apply the same password using Worksheet.Protect(ProtectionType.All) while leaving cell styles and formatting untouched, then saves the file as ProtectedWorksheets.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Add sample data to the first worksheet
            Worksheet sheet0 = workbook.Worksheets[0];
            sheet0.Cells["A1"].PutValue("Sample Data");

            // Add a second worksheet for demonstration
            int secondIndex = workbook.Worksheets.Add();
            Worksheet sheet1 = workbook.Worksheets[secondIndex];
            sheet1.Name = "SecondSheet";
            sheet1.Cells["B2"].PutValue(123);

            // Common password for all worksheets
            string password = "MySecretPwd";

            // Protect each worksheet with the same password while keeping formatting intact
            foreach (Worksheet ws in workbook.Worksheets)
            {
                ws.Protect(ProtectionType.All, password, null);
            }

            // Save the protected workbook
            workbook.Save("ProtectedWorksheets.xlsx");
        }
    }
}
