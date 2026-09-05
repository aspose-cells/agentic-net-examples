// Title: Configure Aspose.Cells column chart data labels to display cell value and series name with a hyphen separator in C#
// AI Prompts: Write C# code using Aspose.Cells to add a column chart and set its first series data labels to show the cell value and the series name separated by a hyphen. | Show how to apply a custom separator string for chart data labels with DataLabelsSeparatorType.Custom in Aspose.Cells. | Create a complete example that builds a workbook, fills sample data, configures the chart labels, and saves the result as an XLSX file.
// Common Searches: asp.net aspocells column chart data label show value and series name with hyphen | c# aspocells custom data label separator for chart series | how to display both value and series name in aspocells chart labels | set datalabelsseparatortype.custom in aspocells example
// Tags: Aspose.Cells chart data label customization | C# set DataLabelsSeparatorType.Custom | Aspose.Cells column chart combined value and series name | Aspose.Cells save workbook as XLSX with chart | Aspose.Cells data labels hyphen separator

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDataLabelsExample
{
    // The program creates a workbook, inserts sample data, adds a column chart, and configures the first series' data labels to show both the cell value and the series name joined by a hyphen before saving the file as DataLabelsValueSeriesHyphen.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Configure data labels to show both value and series name with a hyphen separator
            DataLabels dataLabels = chart.NSeries[0].DataLabels;
            dataLabels.ShowValue = true;                     // Show cell values
            dataLabels.ShowSeriesName = true;                // Show series name
            dataLabels.SeparatorType = DataLabelsSeparatorType.Custom; // Use custom separator
            dataLabels.SeparatorValue = " - ";               // Hyphen with spaces as separator

            // Save the workbook
            workbook.Save("DataLabelsValueSeriesHyphen.xlsx");
        }
    }
}
