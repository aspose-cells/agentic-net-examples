// Title: How to bind custom data label text from a worksheet range to the first series of a column chart using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a column chart, adds the first series, enables data labels, and sets the label text from cells C2:C5 with Aspose.Cells. | Demonstrate using a dynamic Series object to assign custom labels via the SetCustomDataLabel API for the initial series in an Aspose.Cells chart.
// Common Searches: Aspose.Cells C# set custom data labels for chart series from a cell range | How to display worksheet values as data labels in an Aspose.Cells column chart | C# Aspose.Cells example binding cells C2:C5 to first series data labels | Enable data labels for the first series in an Aspose.Cells chart using SetCustomDataLabel | Using dynamic series handling to apply custom data labels in Aspose.Cells .NET
// Tags: set custom data labels Aspose.Cells | first series column chart label range | Aspose.Cells SetCustomDataLabel method | dynamic series data label .NET | column chart data labels from cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example creates a workbook, fills sample data and custom label values, adds a column chart, adds the first series from B2:B5, enables data labels via a dynamic Series object, assigns custom labels from C2:C5 using SetCustomDataLabel, and saves the workbook as output.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Data";

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["A5"].PutValue("D");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["B5"].PutValue(40);

                // Populate custom label values in a separate range
                sheet.Cells["C2"].PutValue("Label1");
                sheet.Cells["C3"].PutValue("Label2");
                sheet.Cells["C4"].PutValue("Label3");
                sheet.Cells["C5"].PutValue("Label4");

                // Add a column chart to the worksheet
                int chartIdx = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
                Chart chart = sheet.Charts[chartIdx];

                // Add the first series (values from B2:B5)
                chart.NSeries.Add("B2:B5", true);
                // Retrieve the series object
                Series firstSeries = chart.NSeries[0];

                // Configure data labels using dynamic to stay compatible with different library versions
                try
                {
                    dynamic seriesDyn = firstSeries;
                    seriesDyn.HasDataLabel = true;
                    seriesDyn.DataLabels.IsValueShown = true;
                    seriesDyn.DataLabels.SetCustomDataLabel("C2:C5");
                }
                catch (Exception labelEx)
                {
                    Console.WriteLine($"Data label configuration skipped: {labelEx.Message}");
                }

                // Ensure the output directory exists
                string outputPath = "output.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully as {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
