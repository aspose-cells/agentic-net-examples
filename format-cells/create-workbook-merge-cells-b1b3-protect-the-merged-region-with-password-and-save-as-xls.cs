// Title: Generate an XLS workbook, merge cells B1:B3, and protect the merged range with a password using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# merge a range of cells and set password protection for the worksheet | How to create an Excel 97-2003 file with merged cells and password protection in .NET | C# code to protect merged cells B1:B3 in an XLS workbook using Aspose.Cells
// Tags: merge cell range B1:B3 Aspose.Cells | worksheet protection with password Aspose.Cells | save as Excel97To2003 XLS Aspose.Cells | merged region password protection .NET | create workbook with merged cells Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, merges cells B1:B3 on the first worksheet, inserts a value, applies password protection to the worksheet, and saves the file as an Excel 97‑2003 (.xls) document.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Merge cells B1:B3 (row index 0‑based, column index 1 for column B)
            // Parameters: startRow, startColumn, totalRows, totalColumns
            sheet.Cells.Merge(0, 1, 3, 1);

            // Optional: put a value into the merged cell
            sheet.Cells["B1"].PutValue("Merged Region");

            // Protect the worksheet with a password.
            // The overload requires oldPassword; pass null when there is no existing password.
            sheet.Protect(ProtectionType.All, "MyPassword", null);

            // Determine output path and ensure directory exists
            string outputPath = "MergedProtected.xls";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as an XLS file
            workbook.Save(outputPath, SaveFormat.Excel97To2003);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
