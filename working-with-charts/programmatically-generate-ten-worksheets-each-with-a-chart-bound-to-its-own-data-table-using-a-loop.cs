// Title: Create ten Excel worksheets each containing a column chart bound to a random data table using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# generate multiple worksheets with individual charts in a loop | how to add a column chart to each sheet programmatically with Aspose.Cells | create ten sheets with random data and bind charts using Aspose.Cells .NET | C# loop to add worksheets and charts to an Excel workbook with Aspose.Cells | save workbook with multiple charted worksheets as .xlsx using Aspose.Cells
// Tags: create worksheets and column charts Aspose.Cells | bind chart to data range Aspose.Cells C# | generate random numeric data Aspose.Cells | configure chart title and legend Aspose.Cells | export workbook to xlsx Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts; // Required for Chart, ChartType, LegendPositionType

// The example creates a new workbook, removes the default sheet, then loops to add ten worksheets named Sheet1‑Sheet10. Each sheet receives a 5‑row by 2‑column table filled with random numbers, a column chart bound to that data, a custom title, and a bottom legend. Finally, the workbook is saved as TenSheetsWithCharts.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (contains one default worksheet)
            Workbook workbook = new Workbook();

            // Remove the default worksheet; we will add our own ten worksheets
            workbook.Worksheets.Clear();

            // Random number generator for sample data
            Random rnd = new Random();

            for (int i = 0; i < 10; i++)
            {
                // Add a new worksheet and give it a meaningful name
                int sheetIndex = workbook.Worksheets.Add();
                Worksheet sheet = workbook.Worksheets[sheetIndex];
                sheet.Name = $"Sheet{i + 1}";

                // Populate a simple data table (5 rows, 2 columns)
                // Column A: Category labels (Item 1 .. Item 5)
                // Column B: Random numeric values
                for (int row = 0; row < 5; row++)
                {
                    sheet.Cells[row, 0].PutValue($"Item {row + 1}");
                    sheet.Cells[row, 1].PutValue(rnd.Next(10, 100));
                }

                // Add a column chart that occupies cells F1 to L15 (adjust as needed)
                int chartUpperLeftRow = 0;      // Row 0 (A1)
                int chartUpperLeftColumn = 5;   // Column F
                int chartLowerRightRow = 15;
                int chartLowerRightColumn = 12; // Column M

                // Add the chart and retrieve its object
                int chartIdx = sheet.Charts.Add(ChartType.Column, chartUpperLeftRow, chartUpperLeftColumn,
                                                chartLowerRightRow, chartLowerRightColumn);
                Chart chart = sheet.Charts[chartIdx];

                // Bind the chart to the data range
                // Values series (vertical) from B1:B5
                chart.NSeries.Add("B1:B5", true);
                // Category (X‑axis) labels from A1:A5
                chart.NSeries.CategoryData = "A1:A5";

                // Optional: set chart title
                chart.Title.Text = $"Chart for {sheet.Name}";
                chart.Title.Font.Size = 12;

                // Optional: set legend position
                chart.Legend.Position = LegendPositionType.Bottom;
            }

            // Save the workbook to a file
            string outputPath = "TenSheetsWithCharts.xlsx";

            // Ensure the directory exists (handle case where outputPath has no directory part)
            string outputDir = Path.GetDirectoryName(outputPath);
            if (string.IsNullOrEmpty(outputDir))
            {
                outputDir = Directory.GetCurrentDirectory();
            }

            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
