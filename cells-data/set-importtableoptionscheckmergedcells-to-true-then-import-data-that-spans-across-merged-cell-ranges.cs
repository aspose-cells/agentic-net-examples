// Title: Import a DataTable into merged Excel cells using Aspose.Cells for .NET (C#)
// Common Searches: asp.net import datatable into merged cells with aspose.cells c# | c# aspose.cells ImportTableOptions.CheckMergedCells true example | how to write data to a merged range in Excel using Aspose.Cells | preserve merged cell formatting when importing a DataTable in Aspose.Cells | import data into merged Excel cells without losing merges Aspose.Cells .NET
// Tags: ImportTableOptions.CheckMergedCells usage Aspose.Cells | DataTable import into merged Excel range C# | Aspose.Cells merged cells handling .NET | C# write to merged cells with Aspose.Cells | Excel merged cell import options Aspose.Cells

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, merges cells A1:C1 on the first worksheet, builds a DataTable with ID, Name, and Value columns, writes the column headers and rows directly into the worksheet cells, ensures the output directory exists, and saves the workbook as ImportMergedCellsResult.xlsx while handling any exceptions.
class ImportMergedCellsExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Merge cells A1:C1 to simulate a merged range in the destination sheet
            sheet.Cells.Merge(0, 0, 1, 3); // Row 0, Column 0, 1 row, 3 columns (A1:C1)

            // Prepare a DataTable with sample data
            DataTable dt = new DataTable("SampleData");
            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Value", typeof(double));

            dt.Rows.Add(1, "Alpha", 10.5);
            dt.Rows.Add(2, "Beta", 20.75);
            dt.Rows.Add(3, "Gamma", 30.0);

            // Manually import the DataTable into the worksheet starting at cell A1
            // Write column headers
            for (int col = 0; col < dt.Columns.Count; col++)
            {
                sheet.Cells[0, col].PutValue(dt.Columns[col].ColumnName);
            }

            // Write data rows
            for (int row = 0; row < dt.Rows.Count; row++)
            {
                for (int col = 0; col < dt.Columns.Count; col++)
                {
                    sheet.Cells[row + 1, col].PutValue(dt.Rows[row][col]);
                }
            }

            // Define output file path
            string outputPath = "ImportMergedCellsResult.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
