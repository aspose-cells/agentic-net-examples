// Title: Merge cells Z4:AA4 and set 45° text rotation in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# merge range Z4:AA4 and rotate text 45 degrees | how to apply text rotation to merged cells with Aspose.Cells .NET | save merged and rotated cells to XLSX using Aspose.Cells | C# code to merge cells Z4 to AA4 and set orientation angle | load workbook with optional password and format cells Aspose.Cells
// Tags: cell range merging Z4 AA4 Aspose.Cells | text orientation 45 degrees Aspose.Cells | XLSX file saving Aspose.Cells | workbook loading with password Aspose.Cells | style application to merged cell Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing workbook (or creates a new one), merges the Z4:AA4 range on the first worksheet, applies a 45‑degree text rotation to the merged cell, and saves the result as an XLSX file, with optional password handling during load.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                // If the file is password‑protected, provide the password here (empty string means no password)
                var loadOptions = new LoadOptions { Password = "" };
                workbook = new Workbook(inputPath, loadOptions);
            }
            else
            {
                workbook = new Workbook();
                // Ensure at least one worksheet is present
                workbook.Worksheets.Clear();
                workbook.Worksheets.Add("Sheet1");
            }

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Merge cells Z4:AA4 (row index 3, columns 25‑26)
            sheet.Cells.Merge(3, 25, 1, 2);

            // Apply 45° rotation to the merged cell (top‑left cell Z4)
            Cell mergedCell = sheet.Cells["Z4"];
            Style style = mergedCell.GetStyle();
            style.RotationAngle = 45;
            mergedCell.SetStyle(style);

            // Save the workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
