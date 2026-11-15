// Title: Set a custom font for chart data labels and let Aspose.Cells automatically resize label shapes in a .NET column chart
// Common Searches: how to set data label font in Aspose.Cells column chart C# | Aspose.Cells automatically adjust data label size after changing font | C# example for customizing chart data label appearance with Aspose.Cells | increase data label font size without clipping in Aspose.Cells chart | apply bold red Arial font to Excel chart labels using Aspose.Cells .NET
// Tags: chart data label font customization Aspose.Cells | auto resize chart label shape .NET | column chart data label styling C# | Aspose.Cells set data label appearance | Excel chart label formatting with Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsCustomDataLabelFont
{
    // This example creates a workbook, adds sample data, inserts a column chart, enables data labels, and customizes the label font to Arial 14pt bold red using Aspose.Cells for .NET. The library automatically resizes the label shapes to accommodate the new styling, and the workbook is saved as CustomDataLabelFont.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["B2"].PutValue(30);
                sheet.Cells["B3"].PutValue(45);
                sheet.Cells["B4"].PutValue(25);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart series
                chart.NSeries.Add("B2:B4", true);
                // Category data can be inferred; explicit setting is optional and omitted for compatibility

                // Enable data labels for the series
                chart.NSeries[0].DataLabels.ShowValue = true;

                // Apply custom font style to data labels
                chart.NSeries[0].DataLabels.Font.Name = "Arial";
                chart.NSeries[0].DataLabels.Font.Size = 14;
                chart.NSeries[0].DataLabels.Font.Color = Color.Red;
                chart.NSeries[0].DataLabels.Font.IsBold = true;
                chart.NSeries[0].DataLabels.Font.IsItalic = false;

                // Save the workbook to a file
                workbook.Save("CustomDataLabelFont.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
