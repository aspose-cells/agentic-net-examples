// Title: Set a triangle‑style data label shape for a column chart using Aspose.Cells in C#
// Common Searches: Aspose.Cells C# set data label shape to triangle for column chart | change chart series data label shape to WedgeRectCallout in Aspose.Cells | C# Aspose.Cells chart data labels custom shape options | how to display triangle‑like data labels in an Aspose.Cells chart | Aspose.Cells DataLabelShapeType enumeration alternatives for triangle shape
// Tags: Aspose.Cells chart data label shape | C# set DataLabelShapeType for series | column chart custom data label appearance | WedgeRectCallout shape Aspose.Cells | triangle data label workaround Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace SetDataLabelShapeType
{
    // The example creates a workbook, adds category and value data, inserts a column chart, enables data labels showing values and categories, and assigns the closest built‑in shape (WedgeRectCallout) to the data labels because the DataLabelShapeType enum does not include a true triangle. The workbook is then saved as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create a new workbook and get the first worksheet.
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // -----------------------------------------------------------------
                // 2. Populate sample data that will be used for the chart.
                // -----------------------------------------------------------------
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Alpha");
                sheet.Cells["A3"].PutValue("Beta");
                sheet.Cells["A4"].PutValue("Gamma");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(15);
                sheet.Cells["B3"].PutValue(30);
                sheet.Cells["B4"].PutValue(45);

                // -----------------------------------------------------------------
                // 3. Add a chart. Aspose.Cells does not expose a dedicated
                //    "custom" chart type enum value, but you can create a chart
                //    of any supported type and then customise its appearance.
                //    Here we use a Column chart as the base.
                // -----------------------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the series.
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // -----------------------------------------------------------------
                // 4. Enable data labels for the first series.
                // -----------------------------------------------------------------
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;          // Show the numeric value.
                series.DataLabels.ShowCategoryName = true;   // Show the category name.

                // -----------------------------------------------------------------
                // 5. Set the shape type of the data labels.
                //    The DataLabelShapeType enum does not contain a "Triangle"
                //    entry. The closest built‑in shape is a "WedgeRectCallout".
                //    If a true triangle is required, a custom drawing layer would
                //    have to be added manually (outside the scope of this demo).
                // -----------------------------------------------------------------
                series.DataLabels.ShapeType = DataLabelShapeType.WedgeRectCallout;

                Console.WriteLine($"DataLabels.ShapeType set to: {series.DataLabels.ShapeType}");

                // -----------------------------------------------------------------
                // 6. Save the workbook.
                // -----------------------------------------------------------------
                string outputPath = "CustomChartWithTriangleDataLabels.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\"");
            }
            catch (Exception ex)
            {
                // Any exception will be written to the console – no failures are hidden.
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
