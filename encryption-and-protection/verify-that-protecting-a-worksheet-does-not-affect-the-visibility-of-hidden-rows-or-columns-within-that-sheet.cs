// Title: Verify that worksheet protection in Aspose.Cells for .NET preserves hidden rows and columns
// Common Searches: does protecting a worksheet with Aspose.Cells change hidden row visibility in C# | keep hidden columns after applying worksheet protection using Aspose.Cells .NET | check IsHidden property after calling Worksheet.Protect in Aspose.Cells | Aspose.Cells protect worksheet without affecting hidden rows example | C# code to hide rows and columns then protect sheet while retaining hidden state
// Tags: worksheet protection preserve hidden rows Aspose.Cells | protect worksheet maintain column hidden state C# | Aspose.Cells IsHidden after Protect method | C# hide rows columns then apply full protection Aspose.Cells | Aspose.Cells worksheet protection all options example

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a workbook, hides a specific row and column, applies full worksheet protection, and then confirms that the IsHidden flags for the hidden row and column remain true before saving the file.
class WorksheetProtectionVisibilityTest
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Hide Row 2 (index 1) and Column B (index 1) using the IsHidden property
            sheet.Cells.Rows[1].IsHidden = true;
            sheet.Cells.Columns[1].IsHidden = true;

            // Protect the worksheet with all protection options
            sheet.Protect(ProtectionType.All);

            // Verify that the hidden states are unchanged after protection
            bool isRow2Hidden = sheet.Cells.Rows[1].IsHidden;
            bool isColumnBHidden = sheet.Cells.Columns[1].IsHidden;

            Console.WriteLine("Row 2 hidden: " + isRow2Hidden);       // Expected: True
            Console.WriteLine("Column B hidden: " + isColumnBHidden); // Expected: True

            // Optional: Save the workbook for manual inspection
            string outputPath = "ProtectedSheet.xlsx";

            // Ensure the directory exists (prevent potential FileNotFoundException)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
