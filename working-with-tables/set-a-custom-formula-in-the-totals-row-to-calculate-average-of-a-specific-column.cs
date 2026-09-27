// Title: Add an AVERAGE formula to the totals row of an Aspose.Cells ListObject table in C#
// AI Prompts: Generate C# code that creates a workbook with Aspose.Cells, inserts a ListObject table, turns on the totals row, and sets the formula AVERAGE([Quantity]) in the totals row for the Quantity column. | Modify an existing Aspose.Cells workbook to enable the totals row of a table and apply a custom structured‑reference formula that computes the average of a specified column. | Write a C# routine that saves an Aspose.Cells workbook to a given file path, ensuring the output directory is created if it does not already exist.
// Common Searches: how to use structured references for totals row formulas with Aspose.Cells ListObject in C# | Aspose.Cells C# set average calculation in table totals row | C# example of adding a totals row to an Excel table and applying AVERAGE formula using Aspose.Cells | create Excel table with totals row and custom formula programmatically with Aspose.Cells
// Tags: Aspose.Cells ListObject totals row formula | C# set AVERAGE structured reference | Excel table average calculation Aspose.Cells | programmatic totals row creation Aspose.Cells | save workbook with directory creation C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;   // Required for ListObject

// Demonstrates creating a workbook, adding a ListObject table, enabling the totals row, labeling it "Average", assigning the structured‑reference formula AVERAGE([Quantity]) to the Quantity column in the totals row, and saving the file while ensuring the output folder exists.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");

            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("Cherry");
            sheet.Cells["B4"].PutValue(30);

            // Define the range for the table (including header and data)
            int firstRow = 0;      // zero‑based index for row 1
            int firstColumn = 0;   // zero‑based index for column A
            int totalRows = 5;     // rows 1‑5 (header + 4 data rows)
            int totalColumns = 2;  // columns A‑B

            // Add a ListObject (table) to the worksheet
            int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn, totalRows, totalColumns, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Enable the totals row
            table.ShowTotals = true;

            // Calculate the index of the totals row (zero‑based)
            int totalsRowIndex = firstRow + totalRows - 1;

            // Set a label in the first column of the totals row
            sheet.Cells[totalsRowIndex, 0].PutValue("Average");

            // Set a formula in the "Quantity" column of the totals row to calculate the average
            // Using structured reference for the column inside the table
            sheet.Cells[totalsRowIndex, 1].Formula = "AVERAGE([Quantity])";

            // Save the workbook to a file
            string outputPath = "TableWithAverageTotals.xlsx";

            // Ensure the directory exists (prevents FileNotFoundException on save)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
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
