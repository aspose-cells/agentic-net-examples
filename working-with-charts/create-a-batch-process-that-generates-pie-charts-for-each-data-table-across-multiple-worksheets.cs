// Title: Batch generate pie charts for each worksheet’s data table in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add a pie chart to each worksheet in a workbook | How to loop through worksheets and create pie charts from columns A and B using Aspose.Cells | Automate batch pie chart creation for multiple sheets in an Excel file with Aspose.Cells .NET | Set chart data range and position for pie charts on every sheet using Aspose.Cells C# | Aspose.Cells generate pie charts for all data tables in an Excel workbook programmatically
// Tags: Aspose.Cells create pie chart per worksheet | C# batch generate Excel charts | set chart data range using CellsHelper | automated pie chart insertion in workbook | iterate Excel worksheets to add charts

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing Excel workbook, iterates through each worksheet, determines the data range in columns A (categories) and B (values), adds a pie chart positioned at rows 5‑20 and columns 0‑7, binds the values and categories, enables data labels, and saves the modified workbook. It includes file‑existence checks and exception handling for robust batch chart generation.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook containing multiple worksheets with data tables
            Workbook workbook = new Workbook(inputPath);

            // Process each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine the last row that contains data in column A (categories)
                int lastRow = sheet.Cells.MaxDataRow;
                if (lastRow < 1) continue; // Skip empty sheets

                // Build range addresses for values (column B) and categories (column A)
                // StartRow = 1 (second row, 0‑based index)
                int startRow = 1;
                int rowsCount = lastRow - startRow + 1; // inclusive count
                if (rowsCount <= 0) continue;

                // Construct address strings in the form "SheetName!A2:B5"
                string valuesStart = CellsHelper.CellIndexToName(startRow, 1); // column B
                string valuesEnd = CellsHelper.CellIndexToName(lastRow, 1);
                string categoriesStart = CellsHelper.CellIndexToName(startRow, 0); // column A
                string categoriesEnd = CellsHelper.CellIndexToName(lastRow, 0);

                string valuesRange = $"{sheet.Name}!{valuesStart}:{valuesEnd}";
                string categoriesRange = $"{sheet.Name}!{categoriesStart}:{categoriesEnd}";

                try
                {
                    // Add a pie chart to the worksheet (positioned at rows 5‑20, columns 0‑7)
                    int chartIdx = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 7);
                    Chart chart = sheet.Charts[chartIdx];

                    // Set a descriptive title
                    chart.Title.Text = $"Pie Chart – {sheet.Name}";

                    // Add the data series (values) and bind the category (X) values
                    int seriesIdx = chart.NSeries.Add(valuesRange, true);
                    chart.NSeries[seriesIdx].XValues = categoriesRange;

                    // Show data labels (values) on the pie slices
                    chart.NSeries[seriesIdx].DataLabels.ShowValue = true;
                }
                catch (Exception chartEx)
                {
                    Console.WriteLine($"Failed to create chart on sheet '{sheet.Name}': {chartEx.Message}");
                }
            }

            // Save the workbook with the newly created charts
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
