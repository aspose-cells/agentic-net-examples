// Title: Create a new Excel workbook, add header and sample rows, and freeze the first column using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that builds a workbook, inserts a header row and ten rows of sample data, then freezes column A. | Show how to use the FreezePanes method to lock the first column while leaving rows scrollable, and save the file as an XLSX. | Provide a complete example that creates a worksheet, populates cells, applies FreezePanes(0,1,…) and writes the workbook to disk.
// Common Searches: Aspose.Cells C# example to freeze column A in an Excel file | how to programmatically add header row and sample data with Aspose.Cells .NET | C# code for creating a new workbook and saving as SampleFreezeColumn.xlsx | using FreezePanes to lock first column while scrolling rows in Aspose.Cells | populate worksheet with numeric data and freeze panes in .NET
// Tags: freeze first column using FreezePanes Aspose.Cells | populate worksheet with sample data C# | create new workbook programmatically Aspose.Cells | save workbook as XLSX Aspose.Cells .NET | add header row Excel C# Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new Workbook, adds a header row and ten rows of sample data to the first worksheet, freezes column A with the FreezePanes method, and saves the result as SampleFreezeColumn.xlsx.
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

            // Add header row
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["C1"].PutValue("Header3");

            // Populate sample data
            for (int i = 2; i <= 10; i++)
            {
                sheet.Cells[i, 0].PutValue("Row " + (i - 1)); // Column A
                sheet.Cells[i, 1].PutValue(i * 10);          // Column B
                sheet.Cells[i, 2].PutValue(i * 100);        // Column C
            }

            // Freeze the first column (Header column)
            // FreezePanes(row, column, totalRows, totalColumns)
            // row = 0 (no rows frozen), column = 1 (freeze column A)
            // totalRows/totalColumns define the visible area; using max data extents
            int totalRows = sheet.Cells.MaxDataRow + 1;
            int totalColumns = sheet.Cells.MaxDataColumn + 1;
            sheet.FreezePanes(0, 1, totalRows, totalColumns);

            // Define output file path
            string outputPath = "SampleFreezeColumn.xlsx";

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
