// Title: How to configure Aspose.Cells chart data labels to show series name and cell value with a hyphen separator in C#
// Common Searches: Aspose.Cells C# set data label text to 'SeriesName - Value' for a column chart | custom separator for chart data labels using Aspose.Cells .NET | display both series name and cell value in Excel chart labels with Aspose.Cells | configure DataLabels.ShowSeriesName and ShowValue with hyphen separator in C# | Aspose.Cells chart label format series name value hyphen
// Tags: Aspose.Cells chart data label customization | DataLabels custom separator Aspose.Cells | C# series name value label format | column chart data labels Aspose.Cells | XLSX chart label hyphen separator

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDataLabelsDemo
{
    // The example creates a workbook, adds sample data and a column chart, then configures the first series' DataLabels to display the series name and cell value concatenated with a hyphen. The workbook is saved as DataLabelsSeriesValueHyphen.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Alpha");
                sheet.Cells["A3"].PutValue("Beta");
                sheet.Cells["A4"].PutValue("Gamma");

                sheet.Cells["B1"].PutValue("Series 1");
                sheet.Cells["B2"].PutValue(15);
                sheet.Cells["B3"].PutValue(30);
                sheet.Cells["B4"].PutValue(45);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the series
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Access the first series
                Series series = chart.NSeries[0];
                series.Name = sheet.Cells["B1"].StringValue; // Set series name explicitly

                // Configure DataLabels to show "SeriesName - Value"
                DataLabels dataLabels = series.DataLabels;
                dataLabels.ShowSeriesName = true;
                dataLabels.ShowValue = true;
                dataLabels.SeparatorType = DataLabelsSeparatorType.Custom;
                dataLabels.SeparatorValue = " - ";
                dataLabels.Position = LabelPositionType.OutsideEnd; // optional positioning
                dataLabels.IsAutoText = true; // let the chart generate the label text

                // Save the workbook
                string outputPath = "DataLabelsSeriesValueHyphen.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
