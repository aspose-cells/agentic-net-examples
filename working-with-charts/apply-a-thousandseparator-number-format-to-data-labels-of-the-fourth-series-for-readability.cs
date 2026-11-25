// Title: How to format data labels with thousand separators for the fourth series in an Aspose.Cells column chart using C#
// Common Searches: C# Aspose.Cells set thousand separator on chart data labels for a specific series | apply number format '#,##0' to fourth series data labels in Aspose.Cells column chart | display commas in column chart data labels using Aspose.Cells .NET API | configure data label position and number format for a series in Aspose.Cells chart
// Tags: Aspose.Cells column chart data label number format | C# set thousand separator for chart series labels | Aspose.Cells specific series data label formatting | Aspose.Cells chart label position outside end | Aspose.Cells apply custom number format to data labels

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDataLabelThousandSeparator
{
    // The example creates a workbook with four data series, adds a column chart, and configures the fourth series to show data labels using the "#,##0" thousand‑separator format and positions the labels outside the column ends before saving the file as XLSX.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for four series
                // Categories in column A
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");
                sheet.Cells["A5"].PutValue("Q4");

                // Series 1 values in column B
                sheet.Cells["B1"].PutValue("Series 1");
                sheet.Cells["B2"].PutValue(12345);
                sheet.Cells["B3"].PutValue(23456);
                sheet.Cells["B4"].PutValue(34567);
                sheet.Cells["B5"].PutValue(45678);

                // Series 2 values in column C
                sheet.Cells["C1"].PutValue("Series 2");
                sheet.Cells["C2"].PutValue(56789);
                sheet.Cells["C3"].PutValue(67890);
                sheet.Cells["C4"].PutValue(78901);
                sheet.Cells["C5"].PutValue(89012);

                // Series 3 values in column D
                sheet.Cells["D1"].PutValue("Series 3");
                sheet.Cells["D2"].PutValue(90123);
                sheet.Cells["D3"].PutValue(101234);
                sheet.Cells["D4"].PutValue(111345);
                sheet.Cells["D5"].PutValue(121456);

                // Series 4 values in column E
                sheet.Cells["E1"].PutValue("Series 4");
                sheet.Cells["E2"].PutValue(131567);
                sheet.Cells["E3"].PutValue(141678);
                sheet.Cells["E4"].PutValue(151789);
                sheet.Cells["E5"].PutValue(161890);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];
                chart.Title.Text = "Sample Chart with Thousand Separator";

                // Add the four series to the chart
                // Series 1
                chart.NSeries.Add("B2:B5", true);
                // Series 2
                chart.NSeries.Add("C2:C5", true);
                // Series 3
                chart.NSeries.Add("D2:D5", true);
                // Series 4
                chart.NSeries.Add("E2:E5", true);

                // Set category data for all series
                chart.NSeries.CategoryData = "A2:A5";

                // Configure data labels for the fourth series (index 3)
                Series fourthSeries = chart.NSeries[3];
                DataLabels dataLabels = fourthSeries.DataLabels;

                // Show the value on the data label
                dataLabels.ShowValue = true;

                // Apply thousand‑separator number format
                // "#,##0" will display numbers like 1,234,567
                dataLabels.NumberFormat = "#,##0";

                // Optional: adjust label position for better readability
                dataLabels.Position = LabelPositionType.OutsideEnd;

                // Save the workbook
                string outputPath = "ChartDataLabelsThousandSeparator.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
