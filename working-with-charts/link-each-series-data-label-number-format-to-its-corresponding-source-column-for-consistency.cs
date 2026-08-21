// Title: Link Chart Series Data Labels to Source Cells with NumberFormatLinked in Aspose.Cells for .NET (C#)
// Description: Creates a workbook, fills raw and formatted columns, adds a column chart, and for each series links data labels to a formatted source range using LinkedSource and NumberFormatLinked, customizes label colors, and saves the file.
// Keywords: Aspose.Cells | C# | .NET | chart data labels | LinkedSource | NumberFormatLinked | column chart | formatted cells | data label font color | Excel automation
// Common Searches: Aspose.Cells link data label to source range | C# set NumberFormatLinked for chart data labels | show formatted values in chart data labels Aspose.Cells | change data label font color in Aspose.Cells chart | link data labels to another worksheet Aspose.Cells
// Developer Intent: The developer needs each series' data label to inherit the number format from its own formatted source column, keeping label display consistent with the source cells.
// Use Cases: Display values such as "100 units" in column‑chart data labels by linking each series to a separate formatted column. | Apply distinct font colors to data labels of multiple series while preserving each series' number format. | Generate an Excel workbook with category labels and two data series, linking each series' labels to different source ranges on the same sheet.
// AI Prompts: Convert the example to a line chart while still linking data label number formats to source cells. | Provide code to link data label number formats to a range on a different worksheet using Aspose.Cells. | Explain how NumberFormatLinked affects data label rendering when the source cells use custom number formats.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, fills raw and formatted columns, adds a column chart, and for each series links data labels to a formatted source range using LinkedSource and NumberFormatLinked, customizes label colors, and saves the file.
    public class LinkDataLabelNumberFormatDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                // Column A – Category
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");

                // Column B – First series values
                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["B2"].PutValue(100);
                sheet.Cells["B3"].PutValue(200);
                sheet.Cells["B4"].PutValue(300);

                // Column C – Second series values
                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["C2"].PutValue(150);
                sheet.Cells["C3"].PutValue(250);
                sheet.Cells["C4"].PutValue(350);

                // Column D – Formatted values for Series1 (e.g., with units)
                sheet.Cells["D1"].PutValue("Formatted1");
                sheet.Cells["D2"].PutValue("100 units");
                sheet.Cells["D3"].PutValue("200 units");
                sheet.Cells["D4"].PutValue("300 units");

                // Column E – Formatted values for Series2
                sheet.Cells["E1"].PutValue("Formatted2");
                sheet.Cells["E2"].PutValue("150 units");
                sheet.Cells["E3"].PutValue("250 units");
                sheet.Cells["E4"].PutValue("350 units");

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // First series (values from B2:B4) linked to D2:D4
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries[0].Name = "Series1";
                chart.NSeries[0].DataLabels.ShowValue = true;
                chart.NSeries[0].DataLabels.LinkedSource = "D2:D4";
                chart.NSeries[0].DataLabels.NumberFormatLinked = true; // link number format
                chart.NSeries[0].DataLabels.Font.Color = Color.Blue;

                // Second series (values from C2:C4) linked to E2:E4
                chart.NSeries.Add("C2:C4", true);
                chart.NSeries[1].Name = "Series2";
                chart.NSeries[1].DataLabels.ShowValue = true;
                chart.NSeries[1].DataLabels.LinkedSource = "E2:E4";
                chart.NSeries[1].DataLabels.NumberFormatLinked = true; // link number format
                chart.NSeries[1].DataLabels.Font.Color = Color.Green;

                // Set category (X) data
                chart.NSeries.CategoryData = "A2:A4";

                // Save the workbook
                workbook.Save("LinkDataLabelNumberFormatDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point required for console application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
