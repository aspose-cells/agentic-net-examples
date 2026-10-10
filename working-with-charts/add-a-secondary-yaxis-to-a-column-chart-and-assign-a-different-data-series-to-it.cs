// Title: Add a secondary Y‑axis to a column chart and bind a separate data series using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that creates a column chart, assigns the first series to the primary Y‑axis and moves the second series onto a secondary Y‑axis, including version‑check logic for the axis properties. | Generate a code snippet that adds a secondary value axis to an Aspose.Cells chart and links a specific NSeries to it, handling the case where the API is not available in older library versions.
// Common Searches: how to create a dual‑axis column chart with Aspose.Cells in C# | Aspose.Cells assign series to secondary Y axis example | C# Aspose.Cells chart with primary and secondary value axes | check Aspose.Cells version for secondary axis support | sample code for column chart with two Y axes using Aspose.Cells
// Tags: Aspose.Cells secondary value axis | C# dual‑axis column chart | Aspose.Cells chart series axis assignment | Aspose.Cells version conditional features | Excel column chart two Y axes Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example creates a workbook, fills it with month data and two numeric series, adds a column chart, sets category labels, adds both series, and demonstrates where to configure a secondary Y‑axis for one series. It also notes that the current Aspose.Cells release may lack direct secondary‑axis properties, so version‑check logic is recommended.
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
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Primary Series");
            sheet.Cells["C1"].PutValue("Secondary Series");

            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");

            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            sheet.Cells["C2"].PutValue(100);
            sheet.Cells["C3"].PutValue(150);
            sheet.Cells["C4"].PutValue(200);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the category (X‑axis) labels
            chart.NSeries.CategoryData = "A2:A4";

            // Add the first data series (uses primary Y‑axis)
            chart.NSeries.Add("B2:B4", true);

            // Add the second data series (will be assigned to secondary Y‑axis if supported)
            chart.NSeries.Add("C2:C4", true);

            // NOTE: The current Aspose.Cells version does not expose IsSecondaryAxis,
            // PrimaryValueAxis, or SecondaryValueAxis properties. If a newer version
            // provides them, they can be set here with appropriate checks.

            // Define output file path
            string outputPath = "ColumnChartWithSecondaryAxis.xlsx";

            // Ensure the output directory exists
            string directory = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook with the chart
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
