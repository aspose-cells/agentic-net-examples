// Title: Add a calculated Total column to an Excel table using structured references and auto‑resize the table with Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that creates a ListObject, adds a 'Total' column using the formula =SalesData[@Quantity]*SalesData[@Price], and saves the workbook. | Demonstrate how to append a new data row and programmatically resize the Aspose.Cells table so the structured‑reference formula propagates to the added row.
// Common Searches: asp.net aspose.cells add calculated column with structured reference formula | how to expand an Excel ListObject range after inserting rows using Aspose.Cells C# | using table structured references in Aspose.Cells to compute column values | auto adjust Excel table size programmatically with Aspose.Cells .NET | create Excel table and total column in C# Aspose.Cells example
// Tags: Aspose.Cells calculated column via table structured reference | C# programmatically resize Excel ListObject | Aspose.Cells create ListObject from cell range | auto‑expand Excel table range with Aspose.Cells | total column using table structured reference .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;   // Required for ListObject

// The sample creates a new workbook, defines a ListObject named 'SalesData' over sample data, adds a 'Total' column whose cells use a structured‑reference formula (=SalesData[@Quantity]*SalesData[@Price]), inserts an extra row, resizes the table to include the new row, and saves the file as StructuredReferenceExample.xlsx.
class StructuredReferenceExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data (Header + 4 rows)
            // Headers: Item, Quantity, Price
            cells["A1"].PutValue("Item");
            cells["B1"].PutValue("Quantity");
            cells["C1"].PutValue("Price");

            // Sample rows
            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue(10);
            cells["C2"].PutValue(0.5);

            cells["A3"].PutValue("Banana");
            cells["B3"].PutValue(5);
            cells["C3"].PutValue(0.3);

            cells["A4"].PutValue("Orange");
            cells["B4"].PutValue(8);
            cells["C4"].PutValue(0.4);

            cells["A5"].PutValue("Grape");
            cells["B5"].PutValue(12);
            cells["C5"].PutValue(0.6);

            // Define the range that will become a table (including headers)
            int firstRow = 0;   // zero‑based index for row 1
            int firstCol = 0;   // column A
            int totalRows = 5;  // header + 4 data rows
            int totalCols = 3;  // Item, Quantity, Price

            // Add a ListObject (Excel Table) over the range
            // The last parameter 'hasHeaders' is true because the first row contains headers
            int tableIndex = sheet.ListObjects.Add(firstRow, firstCol,
                                                   firstRow + totalRows - 1,
                                                   firstCol + totalCols - 1,
                                                   true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "SalesData";
            table.ShowHeaderRow = true;
            // ShowTotalRow property is not available in this version; omitted.

            // Add a new column for Total (Quantity * Price) using a structured reference formula
            int totalColumnIndex = totalCols; // next column after existing ones (zero‑based)

            // Set header for the new column
            cells[0, totalColumnIndex].PutValue("Total");

            // Apply structured reference formula to the new column (excluding header)
            string formula = "=SalesData[@Quantity]*SalesData[@Price]";
            for (int row = 1; row < totalRows; row++)
            {
                cells[row, totalColumnIndex].Formula = formula;
            }

            // Demonstrate automatic range adjustment:
            // Add a new data row below the existing table
            int newRowIndex = firstRow + totalRows; // row index for the new row
            cells[newRowIndex, 0].PutValue("Mango");
            cells[newRowIndex, 1].PutValue(7);
            cells[newRowIndex, 2].PutValue(0.8);

            // Expand the table to include the new row (hasHeaders = true)
            int newTotalRows = newRowIndex - firstRow + 1; // total rows after adding the new row
            table.Resize(firstRow, firstCol, newTotalRows, totalCols, true);

            // Save the workbook
            string outputPath = "StructuredReferenceExample.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
