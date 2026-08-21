// Title: Aspose.Cells for .NET – Enable Table Auto‑Filter and Show Rows with Quantity > Threshold
// Description: Creates a workbook, adds sample data as a ListObject, activates AutoFilter, applies a GreaterThan filter on the Quantity column to keep only rows above a defined limit, and saves the worksheet.
// Keywords: Aspose.Cells | C# auto filter | ListObject filter | custom greater than filter | Excel table filter .NET | threshold filter | Aspose.Cells example | GitHub Aspose.Cells | programmatic Excel auto‑filter
// Common Searches: Aspose.Cells enable auto filter on table C# | filter Excel table rows greater than value using Aspose.Cells | ListObject custom filter threshold Aspose.Cells .NET | how to apply GreaterThan filter to a column with Aspose.Cells | auto‑filter numeric column Aspose.Cells example
// Developer Intent: Create a table, turn on its AutoFilter, and display only rows where the Quantity column exceeds a given threshold.
// Use Cases: Generate an inventory sheet that lists items with stock levels above the reorder point. | Produce a sales report that includes only transactions exceeding a high‑value cutoff. | Export data to Excel while automatically hiding rows whose numeric metric falls below a minimum acceptable value.
// AI Prompts: Write C# code with Aspose.Cells to add a ListObject, enable AutoFilter, and keep rows where column B values are greater than 200. | Show how to resize an existing Aspose.Cells table and apply a custom filter for values less than a specified limit. | Demonstrate setting up multiple AutoFilter criteria (e.g., greater than and less than) on a table column using Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds sample data as a ListObject, activates AutoFilter, applies a GreaterThan filter on the Quantity column to keep only rows above a defined limit, and saves the worksheet.
    public class TableAutoFilterAboveThreshold
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the table (Header + numeric column)
                worksheet.Cells["A1"].PutValue("Item");
                worksheet.Cells["B1"].PutValue("Quantity");
                worksheet.Cells["A2"].PutValue("Apple");
                worksheet.Cells["B2"].PutValue(120);
                worksheet.Cells["A3"].PutValue("Banana");
                worksheet.Cells["B3"].PutValue(80);
                worksheet.Cells["A4"].PutValue("Cherry");
                worksheet.Cells["B4"].PutValue(150);
                worksheet.Cells["A5"].PutValue("Date");
                worksheet.Cells["B5"].PutValue(60);

                // Define the range of the table (rows 0‑4, columns 0‑1)
                int firstRow = 0;
                int firstCol = 0;
                int lastRow = 4;   // zero‑based index of the last data row
                int lastCol = 1;   // zero‑based index of the last column

                // Add a ListObject (table) covering the data range
                int tableIndex = worksheet.ListObjects.Add(firstRow, firstCol, lastRow, lastCol, true);
                ListObject table = worksheet.ListObjects[tableIndex];

                // Enable auto‑filter for the table
                table.HasAutoFilter = true;

                // Ensure the table size includes all rows (optional if range already correct)
                table.Resize(firstRow, firstCol, lastRow, lastCol, true);

                // Define the threshold value
                int threshold = 100;

                // Apply a custom filter on the "Quantity" column (field index 1) to show rows > threshold
                table.AutoFilter.Custom(1, FilterOperatorType.GreaterThan, threshold);
                table.AutoFilter.Refresh();

                // Save the workbook
                string outputPath = Path.Combine(Environment.CurrentDirectory, "TableAutoFilterAboveThreshold.xlsx");
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            TableAutoFilterAboveThreshold.Run();
        }
    }
}
