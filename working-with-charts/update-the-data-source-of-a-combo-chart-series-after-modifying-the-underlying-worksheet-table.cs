// Title: How to update a combo chart’s data source after expanding a ListObject table using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# update combo chart series after adding a row to a ListObject | Refresh chart data range to structured table reference in .NET workbook | Resize Excel table and keep combo chart linked with Aspose.Cells | Set XValues and Values of chart series to table columns using Aspose.Cells C# | Programmatically expand a worksheet table and adjust chart source in C#
// Tags: Aspose.Cells combo chart series data source | ListObject table resize with chart linkage | structured reference XValues Values C# | expand worksheet table adjust chart Aspose | refresh chart after adding table row .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Charts;

// The example loads an existing workbook, locates the "SalesData" ListObject, inserts a new row of data, resizes the table to include the new row, and then updates each series of the first combo chart to use structured references ("=SalesData[Product]" for XValues and "=SalesData[Amount]" for Values) before saving the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Modify the underlying worksheet table
            // -------------------------------------------------
            // Assume the table is named "SalesData".
            ListObject table = sheet.ListObjects["SalesData"];
            if (table == null)
            {
                Console.WriteLine("Table \"SalesData\" not found in the worksheet.");
                return;
            }

            // Determine the first empty row after the current table data.
            int firstDataRow = table.DataRange.FirstRow;
            int dataRowCount = table.DataRange.RowCount;
            int newRowIndex = firstDataRow + dataRowCount; // zero‑based index

            // Insert new data (example values).
            sheet.Cells[newRowIndex, 0].PutValue("New Product"); // Column A – Product
            sheet.Cells[newRowIndex, 1].PutValue(120);          // Column B – Quantity
            sheet.Cells[newRowIndex, 2].PutValue(2500);         // Column C – Amount

            // Expand the table to include the newly added row (hasHeaders = true).
            table.Resize(
                table.DataRange.FirstRow,
                table.DataRange.FirstColumn,
                table.DataRange.RowCount + 1,
                table.DataRange.ColumnCount,
                true);

            // -------------------------------------------------
            // Update the data source of the combo chart series
            // -------------------------------------------------
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found on the worksheet.");
                return;
            }

            Chart comboChart = sheet.Charts[0];

            // Update each series to reference the table columns.
            foreach (Series series in comboChart.NSeries)
            {
                // Category (X‑axis) uses the "Product" column.
                series.XValues = "=SalesData[Product]";
                // Values (Y‑axis) uses the "Amount" column.
                series.Values = "=SalesData[Amount]";
            }

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
