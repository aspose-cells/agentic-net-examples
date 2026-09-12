// Title: How to enable a totals row and set a MEDIAN formula for a column in an Aspose.Cells ListObject table using C#
// AI Prompts: Generate C# code that creates a worksheet, adds a ListObject table, turns on the totals row, labels the totals cell, and applies the formula MEDIAN([Value]) to compute the median of a numeric column. | Write a C# snippet using Aspose.Cells to add a table to a workbook, enable its totals row, and assign a custom median calculation to the totals cell of a specific column. | Provide C# instructions for programmatically inserting a ListObject, showing the totals row, and setting a custom formula that returns the median of the 'Value' column in the totals row.
// Common Searches: Aspose.Cells C# show totals row in ListObject and calculate median | set custom formula for totals row in Aspose.Cells table | C# Aspose.Cells median function in table totals row example | how to add a totals row with median calculation using Aspose.Cells for .NET | programmatically enable totals row and median formula in Excel table with Aspose.Cells
// Tags: Aspose.Cells ListObject enable totals row | Aspose.Cells set totals row formula | C# median calculation in Excel table | Aspose.Cells custom totals row formula | Excel table median using Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables; // Required for ListObject

// Creates a new workbook, adds a ListObject table with sample data, turns on the totals row, labels the totals cell as "Median", assigns the formula MEDIAN([Value]) to compute the median of the "Value" column, and saves the workbook as TableWithMedianTotal.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (Header + numeric values)
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["A5"].PutValue("D");
            sheet.Cells["B5"].PutValue(40);
            sheet.Cells["A6"].PutValue("E");
            sheet.Cells["B6"].PutValue(50);

            // Define the range that will become a table (ListObject)
            int firstRow = 0;        // zero‑based index for row 1
            int firstColumn = 0;     // column A
            int totalRows = 6;       // rows 1‑6
            int totalColumns = 2;    // columns A‑B

            // Add a ListObject (table) to the worksheet
            int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn,
                firstRow + totalRows - 1, firstColumn + totalColumns - 1, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "SampleTable";

            // Enable the Totals row
            table.ShowTotals = true;

            // Calculate the index of the totals row (firstRow + totalRows)
            int totalRowIdx = firstRow + totalRows;

            // Set the label for the totals row in the first column
            sheet.Cells[totalRowIdx, 0].PutValue("Median");

            // Set a custom formula for the "Value" column (index 1) to calculate the median
            sheet.Cells[totalRowIdx, 1].Formula = "MEDIAN([Value])";

            // Save the workbook to a file
            string outputPath = "TableWithMedianTotal.xlsx";

            // Ensure the directory exists (if a directory part is present)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
