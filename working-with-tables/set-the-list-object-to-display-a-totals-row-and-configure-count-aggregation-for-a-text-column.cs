// Title: Add a totals row to an Aspose.Cells ListObject and count non‑empty text entries in a column (C#)
// Common Searches: Aspose.Cells C# how to display a totals row in a ListObject table | C# count non‑empty cells in a text column of an Aspose.Cells ListObject | set totals row label and custom count aggregation for Excel table using Aspose.Cells | example of adding a totals row and calculating row count with Aspose.Cells for .NET | Aspose.Cells ListObject totals row manual count implementation
// Tags: listobject totals row Aspose.Cells C# | count non‑empty text column Aspose.Cells table | add totals row to Excel table using Aspose.Cells | manual aggregation for ListObject column C# | Aspose.Cells calculate row count in table

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables; // Required for ListObject

// Demonstrates creating a workbook, inserting a ListObject with headers, enabling the totals row, labeling the Amount column, manually counting non‑empty Name entries, writing the count to the totals row, and saving the file as ListObjectWithTotals.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (including header row)
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue("Alice");
            sheet.Cells["A3"].PutValue("Bob");
            sheet.Cells["A4"].PutValue("Alice");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["B3"].PutValue(200);
            sheet.Cells["B4"].PutValue(150);

            // Define the range of the list object (zero‑based indices)
            int firstRow = 0;          // Row 1
            int firstColumn = 0;       // Column A
            int totalRows = 4;         // Header + 3 data rows
            int totalColumns = 2;      // Name and Amount columns

            // Add a list object (table) to the worksheet
            int listIndex = sheet.ListObjects.Add(
                firstRow,
                firstColumn,
                firstRow + totalRows - 1,
                firstColumn + totalColumns - 1,
                true); // true => has header row

            ListObject listObject = sheet.ListObjects[listIndex];

            // Enable totals row
            listObject.ShowTotals = true;

            // Set a label for the totals row in the "Amount" column (index 1)
            listObject.ListColumns[1].TotalsRowLabel = "Count:";

            // Manually calculate the count of names and place it in the totals row of the "Name" column (index 0)
            int nameCount = 0;
            for (int r = firstRow + 1; r < firstRow + totalRows; r++) // data rows only
            {
                if (!string.IsNullOrEmpty(sheet.Cells[r, 0].StringValue))
                    nameCount++;
            }
            // Totals row is located just after the last data row
            int totalsRowIndex = firstRow + totalRows;
            sheet.Cells[totalsRowIndex, 0].PutValue(nameCount);

            // Determine output path and ensure the directory exists
            string outputFile = "ListObjectWithTotals.xlsx";
            string outputPath = Path.GetFullPath(outputFile);
            string outputDir = Path.GetDirectoryName(outputPath);

            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
