// Title: Aspose.Cells C# – Add Hyphen‑Separated Series Name and Value to Chart Data Labels
// Description: Creates a workbook, inserts a column chart, and configures the first series’ data labels to show the series name and cell value together, separated by a hyphen (" - "). The labels are placed inside the base of each column before saving the workbook.
// Keywords: aspose.cells | c# chart data labels | .net excel chart | custom separator | show series name | show value | hyphen separator | column chart labels | label position inside base
// Common Searches: Aspose.Cells show series name and value in chart label | C# set custom separator for Aspose.Cells data labels | Aspose.Cells column chart label hyphen | How to display both value and series name in Aspose.Cells chart | Aspose.Cells data label position inside base
// Developer Intent: Configure a chart’s data labels in Aspose.Cells for .NET so that each label combines the series name and its numeric value with a hyphen separator.
// Use Cases: Generate Excel reports where column labels read "Series1 - 10" for quick visual comparison. | Create dashboards that combine identifiers and metrics in a single chart label to reduce clutter. | Apply consistent label formatting across multiple series while positioning labels inside the column base.
// AI Prompts: Write C# code using Aspose.Cells to enable ShowValue and ShowSeriesName on chart data labels, set a custom " - " separator, and place the labels inside the base of each column. | Explain how to change the custom separator to another character or string for chart data labels in Aspose.Cells. | Show how to apply the same hyphen‑separated label configuration to all series in a multi‑series chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDataLabelsExample
{
    // Creates a workbook, inserts a column chart, and configures the first series’ data labels to show the series name and cell value together, separated by a hyphen (" - "). The labels are placed inside the base of each column before saving the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series (values) and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Configure data labels for the first series
            DataLabels dataLabels = chart.NSeries[0].DataLabels;
            dataLabels.ShowValue = true;               // Show cell values
            dataLabels.ShowSeriesName = true;           // Show series name
            dataLabels.SeparatorType = DataLabelsSeparatorType.Custom; // Use custom separator
            dataLabels.SeparatorValue = " - ";          // Hyphen with spaces as separator

            // Optional: position the labels inside the base of each column
            dataLabels.Position = LabelPositionType.InsideBase;

            // Save the workbook
            workbook.Save("DataLabelsSeriesValueHyphen.xlsx");
        }
    }
}
