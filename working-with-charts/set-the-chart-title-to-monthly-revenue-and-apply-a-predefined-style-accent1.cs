using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data for the chart
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Revenue");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["B3"].PutValue(1500);
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B4"].PutValue(1800);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set the chart title text and make it visible
            chart.Title.Text = "Monthly Revenue";
            chart.Title.IsVisible = true;

            // Apply a simple style to the title (replace built‑in style usage)
            chart.Title.Font.Color = Color.DarkBlue;
            chart.Title.Font.IsBold = true;
            chart.Title.Font.Size = 14;
            chart.Title.Font.Name = "Calibri";

            // Define output file path
            string outputPath = "ChartWithTitle.xlsx";

            // Ensure the directory exists (prevents FileNotFoundException on save)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}