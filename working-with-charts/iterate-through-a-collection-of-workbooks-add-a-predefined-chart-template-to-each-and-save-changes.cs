// Title: Add the same column chart to multiple Excel workbooks and overwrite them using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads each .xlsx file from a supplied list, inserts a column chart with a fixed data range on the first worksheet, sets a custom title and legend, and saves the workbook back to its original location using Aspose.Cells. | Create a .NET batch routine that iterates over a collection of workbook paths, applies a predefined column chart template to each workbook's first sheet, and overwrites the existing files with the updated version.
// Common Searches: how to add the same chart to many Excel files using Aspose.Cells C# | batch process Excel workbooks to insert a column chart Aspose.Cells | C# loop to add column chart to multiple .xlsx files and save changes | Aspose.Cells add chart to first worksheet programmatically | overwrite existing Excel workbook after adding chart with Aspose.Cells
// Tags: add column chart to multiple workbooks Aspose.Cells | batch chart insertion .NET Excel | programmatic chart data range Aspose.Cells | overwrite original workbook after chart Aspose.Cells | iterate workbook list Aspose.Cells C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example iterates over a list of Excel file paths, loads each workbook with Aspose.Cells, adds a column chart to the first worksheet using a predefined data range, customizes the chart title and legend, and saves the workbook back to the original file, effectively overwriting it.
class Program
{
    static void Main()
    {
        // Collection of workbook file paths to process
        List<string> workbookPaths = new List<string>
        {
            @"C:\Workbooks\Report1.xlsx",
            @"C:\Workbooks\Report2.xlsx",
            // add more paths as needed
        };

        foreach (string path in workbookPaths)
        {
            // Verify that the file exists before attempting to load
            if (!File.Exists(path))
            {
                Console.WriteLine($"File not found: {path}");
                continue;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(path);

                // Get the first worksheet (or specify the target worksheet)
                Worksheet sheet = workbook.Worksheets[0];

                // Define the position and size of the chart (in cells)
                int upperLeftRow = 5;      // row index where chart starts
                int upperLeftColumn = 1;   // column index where chart starts
                int height = 15;           // number of rows the chart occupies
                int width = 10;            // number of columns the chart occupies

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, upperLeftRow, upperLeftColumn, height, width);
                Chart chart = sheet.Charts[chartIndex];

                // Set chart title
                chart.Title.Text = "Sales Summary";

                // Example data range for the chart (adjust to your data layout)
                // Assuming data is in A1:B5 (A column = categories, B column = values)
                int dataFirstRow = 0;
                int dataFirstColumn = 0;
                int dataLastRow = 4;
                int dataLastColumn = 1;

                // Build the range string (e.g., Sheet1!$A$1:$B$5)
                string range = $"{sheet.Name}!${CellsHelper.ColumnIndexToName(dataFirstColumn)}${dataFirstRow + 1}:${CellsHelper.ColumnIndexToName(dataLastColumn)}${dataLastRow + 1}";

                // Add a series to the chart using the data range
                int seriesIndex = chart.NSeries.Add(range, true);
                chart.NSeries[seriesIndex].Name = "Quarterly Sales";

                // Optional: customize axes, legend, etc.
                chart.ShowLegend = true;
                chart.Legend.Position = LegendPositionType.Right;

                // Save changes back to the original file
                workbook.Save(path, SaveFormat.Xlsx);
                Console.WriteLine($"Chart added and workbook saved: {path}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{path}': {ex.Message}");
            }
        }
    }
}
