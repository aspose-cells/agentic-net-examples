// Title: How to hide data labels for the third series in a column chart using Aspose.Cells for C#
// AI Prompts: Write C# code that creates a column chart with three series in Aspose.Cells, enables data labels for all series, then disables the labels for the third series. | Show how to use the Series.DataLabels.IsDeleted property in Aspose.Cells to remove data labels from a specific chart series in C#. | Provide a step‑by‑step example of toggling data label visibility per series in an Excel workbook generated with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# hide data labels for only one series in a column chart | How to turn off data labels for the third series in an Excel chart using Aspose.Cells .NET | C# Aspose.Cells chart series label visibility control example
// Tags: Aspose.Cells series data label removal C# | column chart label display Aspose.Cells | disable third series labels Aspose.Cells | C# Excel chart label customization Aspose.Cells | Series.DataLabels.IsDeleted usage Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example creates a new workbook, populates three data series, adds a column chart, enables data labels for all series, then disables the data labels for the third series by setting its Series.DataLabels.IsDeleted property, and finally saves the workbook as an XLSX file.
    public class DisableThirdSeriesDataLabels
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for three series
                // Category column
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");

                // Series 1 values
                sheet.Cells["B1"].PutValue("Series 1");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Series 2 values
                sheet.Cells["C1"].PutValue("Series 2");
                sheet.Cells["C2"].PutValue(15);
                sheet.Cells["C3"].PutValue(25);
                sheet.Cells["C4"].PutValue(35);

                // Series 3 values
                sheet.Cells["D1"].PutValue("Series 3");
                sheet.Cells["D2"].PutValue(12);
                sheet.Cells["D3"].PutValue(22);
                sheet.Cells["D4"].PutValue(32);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Add the three series to the chart
                chart.NSeries.Add("B2:B4", true); // Series 1
                chart.NSeries.Add("C2:C4", true); // Series 2
                chart.NSeries.Add("D2:D4", true); // Series 3

                // Set category (X) data
                chart.NSeries.CategoryData = "A2:A4";

                // Enable data labels for all series (optional)
                foreach (Series s in chart.NSeries)
                {
                    s.DataLabels.ShowValue = true;
                }

                // Disable data labels for the third series (index 2)
                Series thirdSeries = chart.NSeries[2];
                // Hide all data labels for this series
                thirdSeries.DataLabels.IsDeleted = true;

                // Save the workbook
                string outputPath = "DisableThirdSeriesDataLabels.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
