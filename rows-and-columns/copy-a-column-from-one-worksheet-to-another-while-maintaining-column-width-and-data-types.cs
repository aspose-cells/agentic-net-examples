// Title: Copy a column between worksheets while preserving original width, data types, and formatting using Aspose.Cells for .NET
// Common Searches: Aspose.Cells copy column from Sheet1 to Sheet2 while keeping original column width in C# | C# copy entire Excel column with original data types and formatting using Aspose.Cells | How to set destination column width equal to source column with Aspose.Cells .NET | Iterate rows to duplicate column values and styles between worksheets in Aspose.Cells | Copy column B to column D in an Excel workbook programmatically with Aspose.Cells for .NET
// Tags: copy column between worksheets Aspose.Cells | preserve column width Aspose.Cells .NET | copy cell values with data types Aspose.Cells | transfer cell style Aspose.Cells | set column width from source Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook, ensures the source and destination worksheets exist, copies all cells from a source column to a target column row by row, preserving each cell's value type and style, and matches the destination column width to the source before saving the updated file.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source workbook
            string sourcePath = "source.xlsx";

            // Ensure the source file exists; create a minimal workbook if it does not
            if (!File.Exists(sourcePath))
            {
                Workbook tempWb = new Workbook();
                tempWb.Worksheets.Add("Sheet1");
                tempWb.Worksheets.Add("Sheet2");
                // Optionally add sample data to Sheet1 column B
                tempWb.Worksheets["Sheet1"].Cells["B1"].PutValue("Sample");
                tempWb.Save(sourcePath);
            }

            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // Get or create the source worksheet
            Worksheet sourceSheet = workbook.Worksheets["Sheet1"];
            if (sourceSheet == null)
            {
                sourceSheet = workbook.Worksheets.Add("Sheet1");
            }

            // Get or create the destination worksheet
            Worksheet destinationSheet = workbook.Worksheets["Sheet2"];
            if (destinationSheet == null)
            {
                destinationSheet = workbook.Worksheets.Add("Sheet2");
            }

            // Define source and destination column indexes (0‑based)
            int sourceColumnIndex = 1;      // Column B
            int destinationColumnIndex = 3; // Column D

            // Preserve column width
            double sourceColumnWidth = sourceSheet.Cells.GetColumnWidth(sourceColumnIndex);
            destinationSheet.Cells.SetColumnWidth(destinationColumnIndex, sourceColumnWidth);

            // Determine the last row with data in the source column
            int lastRow = sourceSheet.Cells.MaxDataRow;

            // Copy each cell (value, data type, style)
            for (int row = 0; row <= lastRow; row++)
            {
                Cell srcCell = sourceSheet.Cells[row, sourceColumnIndex];
                Cell destCell = destinationSheet.Cells[row, destinationColumnIndex];

                // Copy the cell value (preserves the underlying data type)
                destCell.PutValue(srcCell.Value);

                // Copy the cell style (number format, font, alignment, etc.)
                destCell.SetStyle(srcCell.GetStyle());
            }

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
