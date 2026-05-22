using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class UpdateWaterfallCharts
{
    static void Main()
    {
        // Load the workbook that contains the imported data and the Waterfall charts.
        Workbook workbook = new Workbook("input.xlsx");

        // Define the new data range that the Waterfall charts should use.
        // Adjust the sheet name and range as needed after your bulk import.
        string newDataRange = "Data!$A$2:$B$10"; // Column A = categories, Column B = values

        // Iterate through every worksheet in the workbook.
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through each chart on the current worksheet.
            for (int c = 0; c < sheet.Charts.Count; c++)
            {
                Chart chart = sheet.Charts[c];

                // Process only Waterfall charts.
                if (chart.Type == ChartType.Waterfall)
                {
                    // Remove any existing series.
                    chart.NSeries.Clear();

                    // Split the range into category (A column) and values (B column).
                    string[] parts = newDataRange.Split(':');
                    string categoryRange = parts[0]; // e.g., "Data!$A$2"
                    string valueRange = parts[1];    // e.g., "$B$10"

                    // Build full address for the values series.
                    // The values part must include the sheet name as well.
                    string valuesAddress = $"Data!${valueRange}";

                    // Add the new values series (vertical orientation).
                    chart.NSeries.Add(valuesAddress, true);

                    // Set the category (X‑axis) data.
                    chart.NSeries.CategoryData = $"{categoryRange}:{categoryRange.Replace("$A$", "$A$10")}";

                    // Recalculate the chart so that it reflects the new data.
                    chart.Calculate(new ChartCalculateOptions { UpdateAllPoints = true });
                }
            }
        }

        // Save the updated workbook.
        workbook.Save("output.xlsx");
    }
}