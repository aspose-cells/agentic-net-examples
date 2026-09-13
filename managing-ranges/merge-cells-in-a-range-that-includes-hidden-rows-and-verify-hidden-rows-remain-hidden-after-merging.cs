// Title: Merge a cell range that includes hidden rows while preserving the hidden state using Aspose.Cells for .NET
// AI Prompts: Hide a specific row, merge a vertical range that spans the hidden row with Aspose.Cells, then read the IsHidden flag to confirm it stayed hidden before saving. | Programmatically merge cells A2:A4 in a worksheet, ensure the hidden row remains hidden, and export the workbook to an XLSX file using the Aspose.Cells .NET API. | Create a new workbook, set row 3 as hidden, apply Cells.Merge on rows 2‑4, validate the hidden property, and write the result to disk.
// Common Searches: Aspose.Cells .NET keep hidden rows after merging a cell range | how to verify hidden row stays hidden when merging cells with Aspose.Cells | C# merge range that contains a hidden row and check IsHidden property | preserve row hidden state during Cells.Merge operation in Aspose.Cells
// Tags: Aspose.Cells merge range over hidden rows | C# hide row then merge cells | verify hidden row after merge | save workbook with merged hidden rows to XLSX | Cells.Merge hidden row handling

using Aspose.Cells;
using System;
using System.IO;

// The example creates a workbook, hides row 3, merges cells A2:A4 (which includes the hidden row), checks that the hidden flag of row 3 remains true, and saves the file as MergedHiddenRows.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Fill some sample data
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue("Row1");
            sheet.Cells["A3"].PutValue("Row2");
            sheet.Cells["A4"].PutValue("Row3");

            // Hide row 3 (zero‑based index 2)
            sheet.Cells.Rows[2].IsHidden = true;

            // Merge cells A2:A4 (range includes the hidden row)
            // Parameters: startRow, startColumn, totalRows, totalColumns
            sheet.Cells.Merge(1, 0, 3, 1);

            // Verify that the hidden row remains hidden after merging
            bool isRow3StillHidden = sheet.Cells.Rows[2].IsHidden;
            Console.WriteLine("Row 3 hidden after merge: " + isRow3StillHidden);

            // Define output file path
            string outputPath = "MergedHiddenRows.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
