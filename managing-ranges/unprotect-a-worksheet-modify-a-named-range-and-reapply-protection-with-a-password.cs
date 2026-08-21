// Title: Aspose.Cells for .NET – Unprotect Worksheet, Edit Named Range, and Re‑protect with a New Password
// Description: Load or create an Excel file, unprotect the first sheet using the current password, retrieve the named range "MyRange", change its first cell, then protect the sheet again with a new password and save the workbook.
// Keywords: Aspose.Cells unprotect worksheet | modify named range C# | protect worksheet with password .NET | update cell in named range programmatically | Excel sheet protection Aspose.Cells | C# Aspose.Cells workbook security
// Common Searches: how to unprotect an Excel sheet with Aspose.Cells | change value in a named range after unprotecting sheet | re‑apply worksheet protection with a different password using Aspose.Cells | Aspose.Cells example edit protected named range | C# code to unprotect, modify, and protect Excel worksheet
// Developer Intent: Remove sheet protection, update a cell in a named range, and apply new protection with a different password.
// Use Cases: Automate data refresh in a specific named range of a secured workbook. | Implement password‑policy changes after programmatic edits to protected sheets. | Integrate Excel updates into a CI/CD pipeline without manual unprotecting.
// AI Prompts: Write C# code using Aspose.Cells to unprotect a worksheet with a known password, modify the first cell of a named range, and protect the sheet again with a new password. | Show how to safely retrieve and edit a named range after unprotecting a sheet, then re‑apply protection in Aspose.Cells for .NET. | Provide an Aspose.Cells example that creates a workbook with a named range, protects it, and demonstrates the unprotect‑modify‑protect workflow.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Load or create an Excel file, unprotect the first sheet using the current password, retrieve the named range "MyRange", change its first cell, then protect the sheet again with a new password and save the workbook.
    public class UnprotectModifyProtectDemo
    {
        // Entry point required for the console application
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists; if not, create a minimal workbook with a named range
            if (!File.Exists(inputPath))
            {
                var wb = new Workbook();
                var ws = wb.Worksheets[0];
                ws.Name = "Sheet1";

                // Populate some data
                ws.Cells["A1"].PutValue("Original Value");

                // Create a named range "MyRange" covering A1
                int index = wb.Worksheets.Names.Add("MyRange");
                wb.Worksheets.Names[index].RefersTo = "Sheet1!$A$1";

                // Protect the sheet with the old password for demonstration
                ws.Protect(ProtectionType.All, "oldPassword", null);

                wb.Save(inputPath);
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Unprotect the worksheet using the existing password
            worksheet.Unprotect("oldPassword");

            // Retrieve the named range "MyRange"
            Name namedRange = workbook.Worksheets.Names["MyRange"];
            if (namedRange != null)
            {
                // Get the actual range object (fully qualified to avoid ambiguity)
                Aspose.Cells.Range range = namedRange.GetRange();

                // Modify the first cell of the range if it exists
                if (range != null && range.RowCount > 0 && range.ColumnCount > 0)
                {
                    range[0, 0].PutValue("Modified Value");
                }
            }

            // Re‑apply protection with a new password
            worksheet.Protect(ProtectionType.All, "newPassword", null);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}
