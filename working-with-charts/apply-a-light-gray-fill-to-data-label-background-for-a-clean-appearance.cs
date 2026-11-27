// Title: Apply a light gray background fill to chart data labels using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set light gray fill for chart data label background | How to change data label area color in a column chart with Aspose.Cells | Add a thin border to data labels in an Aspose.Cells chart using C# | Customize data label appearance (fill and border) in Aspose.Cells .NET | Set solid fill pattern for chart data labels in an Aspose.Cells workbook
// Tags: chart data label fill format Aspose.Cells | solid fill pattern for data label area .NET | data label border styling Aspose.Cells | column chart data label customization C# | data label area color configuration Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsDataLabelBackground
{
    // The example creates a workbook, adds a column chart with sample sales data, enables value display on the first series' data labels, applies a solid light gray fill to the data label background, adds a visible dark gray thin border, and saves the result as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");
                sheet.Cells["A5"].PutValue("Q4");

                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["B4"].PutValue(180);
                sheet.Cells["B5"].PutValue(200);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Access the DataLabels of the first series
                DataLabels dataLabels = chart.NSeries[0].DataLabels;

                // Show the value on each data label
                dataLabels.ShowValue = true;
                dataLabels.Position = LabelPositionType.InsideEnd;

                // Apply a light gray fill to the background of the data labels
                dataLabels.Area.FillFormat.Pattern = FillPattern.Solid;
                dataLabels.Area.FillFormat.SolidFill.Color = Color.LightGray;

                // Optional: set a thin border so the label stands out
                dataLabels.Border.IsVisible = true;
                dataLabels.Border.Color = Color.DarkGray;
                dataLabels.Border.Weight = WeightType.SingleLine; // thin border

                // Save the workbook
                string outputPath = "DataLabelsLightGrayBackground.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
