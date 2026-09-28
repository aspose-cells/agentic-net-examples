// Title: How to refresh the data source of a combo chart after expanding a ListObject table using Aspose.Cells for .NET
// AI Prompts: Write C# code that adds a new row to a ListObject, resizes the table, and updates the XValues and Values of each series in the first combo chart to reference the expanded range with Aspose.Cells. | Generate a method that iterates through all series of a combo chart and sets their data ranges based on a resized worksheet table using the Aspose.Cells API. | Create a routine that programmatically synchronizes a combo chart’s series with a dynamically growing Excel table, handling table resize and chart series range updates in C#.
// Common Searches: aspnet update combo chart series after adding rows to a ListObject table | c# Aspose.Cells refresh chart data range when table size changes | how to programmatically adjust XValues for combo chart in Aspose.Cells | expand Excel table and keep combo chart linked using Aspose.Cells for .NET | dynamic chart source update with Aspose.Cells ListObject resize
// Tags: combo chart series range update Aspose.Cells | resize ListObject table C# | chart data source synchronization Aspose.Cells | dynamic Excel chart range Aspose.Cells | programmatic chart refresh after table expansion

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Tables;

// The example loads an existing workbook, adds a new row to a ListObject named "SalesTable", expands the table to include the new row, then loops through each series of the first combo chart on the worksheet, rebuilding the X‑axis and Y‑axis range strings to point to the enlarged table, and finally saves the updated workbook.
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
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook.
            Workbook workbook = new Workbook(inputPath);

            // Get the worksheet that contains the table and the combo chart.
            Worksheet sheet = workbook.Worksheets["Data"]; // adjust sheet name as needed
            if (sheet == null)
            {
                Console.WriteLine("Error: Worksheet 'Data' not found.");
                return;
            }

            // Retrieve the table (ListObject) by its name.
            ListObject table = sheet.ListObjects["SalesTable"]; // adjust table name as needed
            if (table == null)
            {
                Console.WriteLine("Error: Table 'SalesTable' not found.");
                return;
            }

            // -------------------- Modify the underlying table --------------------
            // Example: add a new data row to the table.
            int newRowIndex = table.DataRange.FirstRow + table.DataRange.RowCount; // first empty row after current data
            sheet.Cells[newRowIndex, table.DataRange.FirstColumn].PutValue(DateTime.Today); // Date column (assumed first column)
            sheet.Cells[newRowIndex, table.DataRange.FirstColumn + 1].PutValue(1500);        // Sales column
            sheet.Cells[newRowIndex, table.DataRange.FirstColumn + 2].PutValue(300);         // Quantity column

            // Expand the table to include the newly added row.
            // The last parameter indicates whether the table has headers (true for typical tables).
            table.Resize(
                table.DataRange.FirstRow,
                table.DataRange.FirstColumn,
                table.DataRange.RowCount + 1,
                table.DataRange.ColumnCount,
                true);

            // -------------------- Update the combo chart series --------------------
            // Assume the combo chart is the first chart on the worksheet.
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found on the worksheet.");
                return;
            }

            Chart comboChart = sheet.Charts[0];

            // Update each series to point to the expanded table range.
            for (int i = 0; i < comboChart.NSeries.Count; i++)
            {
                Series series = comboChart.NSeries[i];

                // Determine which data column this series uses.
                // Here we map series index 0 -> column B, 1 -> column C, etc.
                int dataColumnIndex = table.DataRange.FirstColumn + 1 + i; // offset by 1 to skip the X‑axis column

                string sheetName = sheet.Name;

                // X‑axis range (e.g., Data!A2:A6)
                string xRange = $"{sheetName}!{CellsHelper.ColumnIndexToName(table.DataRange.FirstColumn)}{table.DataRange.FirstRow + 1}:{CellsHelper.ColumnIndexToName(table.DataRange.FirstColumn)}{table.DataRange.FirstRow + table.DataRange.RowCount}";

                // Y‑axis range for this series (e.g., Data!B2:B6)
                string yRange = $"{sheetName}!{CellsHelper.ColumnIndexToName(dataColumnIndex)}{table.DataRange.FirstRow + 1}:{CellsHelper.ColumnIndexToName(dataColumnIndex)}{table.DataRange.FirstRow + table.DataRange.RowCount}";

                // Apply the new ranges to the series.
                series.XValues = xRange;
                series.Values = yRange;
            }

            // Save the workbook with the updated chart.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
