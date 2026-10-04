// Title: Generate a workbook with ten worksheets, each containing its own column chart bound to a data table using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a Workbook, loops to add ten worksheets, populates a small data table on each sheet, and inserts a column chart linked to that table with Aspose.Cells. | Produce a C# example that assigns a unique chart title and series name per worksheet while positioning the chart below the data range in a ten‑sheet workbook using Aspose.Cells.
// Common Searches: how to add a column chart to each worksheet in a loop with Aspose.Cells C# | create multiple sheets with individual charts using Aspose.Cells for .NET | Aspose.Cells generate ten worksheets each with its own chart programmatically | C# loop to add worksheets and bind charts to data ranges in an Excel file | set chart title per worksheet Aspose.Cells example
// Tags: Aspose.Cells create worksheets with column charts | C# loop add chart to each sheet | bind chart series to data range Aspose.Cells | generate Excel workbook with multiple charts .NET | set chart title per worksheet Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a Workbook, iterates from 1 to 10 (using the default sheet for the first iteration and adding new sheets for the rest), fills cells A1:A5 with item labels and B1:B5 with values offset by the sheet index, adds a column chart positioned below the data, sets a sheet‑specific title and series name, and saves the file as TenSheetsWithCharts.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (contains one default worksheet)
            Workbook workbook = new Workbook();

            // Generate 10 worksheets, each with its own data table and chart
            for (int i = 1; i <= 10; i++)
            {
                Worksheet sheet;

                // Use the default worksheet for the first iteration, otherwise add a new one
                if (i == 1)
                {
                    sheet = workbook.Worksheets[0];
                    sheet.Name = $"Sheet{i}";
                }
                else
                {
                    sheet = workbook.Worksheets.Add($"Sheet{i}");
                }

                // Populate sample data in columns A (categories) and B (values) for rows 1‑5
                for (int row = 0; row < 5; row++)
                {
                    sheet.Cells[row, 0].PutValue($"Item {row + 1}"); // Column A
                    sheet.Cells[row, 1].PutValue(row + i);          // Column B (different per sheet)
                }

                // Add a column chart positioned below the data table
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 7);
                Chart chart = sheet.Charts[chartIndex];

                // Set chart title
                chart.Title.Text = $"Chart for Sheet{i}";

                // Add a series: values from B1:B5 (categories are taken automatically)
                int seriesIndex = chart.NSeries.Add("B1:B5", true);
                // Set series name
                chart.NSeries[seriesIndex].Name = $"Series {i}";
            }

            // Save the workbook to a file
            string outputPath = "TenSheetsWithCharts.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
