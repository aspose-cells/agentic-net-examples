// Title: How to create a reusable chart template with custom data label formatting and apply it to other charts using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# copy data label settings from one chart to another | Create chart template with predefined data labels in Aspose.Cells .NET | Apply custom data label font and number format to a new column chart using Aspose.Cells | Reuse chart style template across multiple worksheets with Aspose.Cells | C# Aspose.Cells example for copying DataLabels properties between charts
// Tags: chart template data label configuration Aspose.Cells | copy data label properties C# Aspose.Cells | custom data label formatting column chart .NET | reuse chart style template Excel Aspose.Cells | apply data label settings to new chart Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ChartTemplateDemo
{
    // The example creates an Excel workbook, defines a template worksheet with a column chart whose data label properties (value, category name, position, font style, background, separator, number format) are set, then adds another worksheet with its own data and a new column chart, copies the template's DataLabels settings to the new chart via a helper method, prints the applied settings, and saves the file as XLSX.
    class Program
    {
        static void Main(string[] args)
        {
            const string outputPath = "ChartTemplateDemo.xlsx";

            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // 1. Create a worksheet that will hold the chart template
                // -------------------------------------------------
                Worksheet templateSheet = workbook.Worksheets[workbook.Worksheets.Add()];
                templateSheet.Name = "Template";

                // Sample data for the template chart
                templateSheet.Cells["A1"].PutValue("Category");
                templateSheet.Cells["B1"].PutValue("Value");
                templateSheet.Cells["A2"].PutValue("A");
                templateSheet.Cells["A3"].PutValue("B");
                templateSheet.Cells["A4"].PutValue("C");
                templateSheet.Cells["B2"].PutValue(10);
                templateSheet.Cells["B3"].PutValue(20);
                templateSheet.Cells["B4"].PutValue(30);

                // Add a column chart that will serve as the template
                int templateChartIndex = templateSheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart templateChart = templateSheet.Charts[templateChartIndex];
                templateChart.NSeries.Add("B2:B4", true);
                templateChart.NSeries.CategoryData = "A2:A4";

                // -------------------------------------------------
                // 2. Configure DataLabels on the template chart
                // -------------------------------------------------
                Series templateSeries = templateChart.NSeries[0];
                DataLabels templateDl = templateSeries.DataLabels;

                templateDl.ShowValue = true;
                templateDl.ShowCategoryName = true;
                templateDl.Position = LabelPositionType.InsideEnd;
                templateDl.Font.Color = Color.DarkBlue;
                templateDl.Font.Size = 10;
                templateDl.Font.IsBold = true;
                templateDl.BackgroundMode = BackgroundMode.Transparent;
                templateDl.SeparatorType = DataLabelsSeparatorType.Comma;
                templateDl.NumberFormat = "0.00";
                templateDl.ApplyFont();

                // -------------------------------------------------
                // 3. Create a worksheet that will receive a chart based on the template
                // -------------------------------------------------
                Worksheet dataSheet = workbook.Worksheets[workbook.Worksheets.Add()];
                dataSheet.Name = "Data1";

                // Different data for the new chart
                dataSheet.Cells["A1"].PutValue("Category");
                dataSheet.Cells["B1"].PutValue("Value");
                dataSheet.Cells["A2"].PutValue("X");
                dataSheet.Cells["A3"].PutValue("Y");
                dataSheet.Cells["A4"].PutValue("Z");
                dataSheet.Cells["B2"].PutValue(15);
                dataSheet.Cells["B3"].PutValue(25);
                dataSheet.Cells["B4"].PutValue(35);

                // -------------------------------------------------
                // 4. Create a new chart on the data sheet and apply the template settings
                // -------------------------------------------------
                int newChartIndex = dataSheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart newChart = dataSheet.Charts[newChartIndex];

                // Bind the new chart to its data
                newChart.NSeries.Add("B2:B4", true);
                newChart.NSeries.CategoryData = "A2:A4";

                // Copy DataLabels settings from the template series
                CopyDataLabelSettings(templateDl, newChart.NSeries[0].DataLabels);

                // -------------------------------------------------
                // 5. Output a few DataLabels properties to verify they were applied
                // -------------------------------------------------
                DataLabels newDl = newChart.NSeries[0].DataLabels;
                Console.WriteLine("DataLabels on the new chart:");
                Console.WriteLine($"  ShowValue: {newDl.ShowValue}");
                Console.WriteLine($"  ShowCategoryName: {newDl.ShowCategoryName}");
                Console.WriteLine($"  Position: {newDl.Position}");
                Console.WriteLine($"  Font Color: {newDl.Font.Color}");
                Console.WriteLine($"  NumberFormat: {newDl.NumberFormat}");

                // -------------------------------------------------
                // 6. Save the workbook
                // -------------------------------------------------
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        private static void CopyDataLabelSettings(DataLabels source, DataLabels target)
        {
            target.ShowValue = source.ShowValue;
            target.ShowCategoryName = source.ShowCategoryName;
            target.Position = source.Position;
            target.Font.Color = source.Font.Color;
            target.Font.Size = source.Font.Size;
            target.Font.IsBold = source.Font.IsBold;
            target.BackgroundMode = source.BackgroundMode;
            target.SeparatorType = source.SeparatorType;
            target.NumberFormat = source.NumberFormat;
            target.ApplyFont();
        }
    }
}
