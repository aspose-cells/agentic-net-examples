// Title: C# – Apply Light Gray Background to Chart Data Labels with Aspose.Cells
// Description: Creates a workbook, adds sample data, inserts a column chart, enables data labels, and sets a solid light‑gray fill for the label background before saving the file.
// Keywords: Aspose.Cells data label background | C# chart data label fill | light gray data label Aspose | Excel chart formatting .NET | solid fill pattern Aspose.Cells | chart label styling C# | US developers Aspose.Cells | UK Aspose.Cells examples
// Common Searches: Aspose.Cells change data label background color | C# set solid fill for chart data labels | light gray background for Excel chart labels | how to format data labels in Aspose.Cells | chart label styling .NET
// Developer Intent: Add a solid light‑gray fill to the background of chart data labels using Aspose.Cells for .NET.
// Use Cases: Generate a column chart where each data label has a light‑gray background to improve readability. | Apply a consistent label style across multiple charts in an automated Excel report. | Retrofitting existing workbooks with corporate‑approved label colors without manual editing.
// AI Prompts: Show C# code that sets a solid light gray fill for chart data labels with Aspose.Cells. | Give an example of applying a light gray background to pie‑chart data labels in .NET. | Explain how to customize data label fill pattern, background color, and border in Aspose.Cells charts.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds sample data, inserts a column chart, enables data labels, and sets a solid light‑gray fill for the label background before saving the file.
    public class DataLabelsLightGrayBackgroundDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["B4"].PutValue(30);

                // Add a column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
                Chart chart = worksheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable data labels for the first series
                DataLabels dataLabels = chart.NSeries[0].DataLabels;
                dataLabels.ShowValue = true;

                // Apply a solid light gray fill to the data label background
                dataLabels.Area.FillFormat.Pattern = FillPattern.Solid;
                dataLabels.Area.BackgroundColor = Color.LightGray;

                // Save the workbook
                string outputPath = "DataLabelsLightGrayBackgroundDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
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
            DataLabelsLightGrayBackgroundDemo.Run();
        }
    }
}
