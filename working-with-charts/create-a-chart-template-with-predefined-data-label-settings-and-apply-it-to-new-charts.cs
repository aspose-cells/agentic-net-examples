// Title: Load a .crtx chart template and customize data label appearance in Aspose.Cells for .NET
// AI Prompts: Write C# code that reads a .crtx file into a byte array, creates a column chart, applies the template with ChangeTemplate, and sets the data labels to show values with a dark‑blue font. | Show C# how to add a chart using the Add(byte[], string, bool, int, int, int, int) overload with a template byte array, then change the first series data label position to InsideEnd. | Create a reusable C# method that takes a template path and a workbook, adds a chart based on the template, and applies custom data label settings such as font color and label position.
// Common Searches: how to apply a .crtx chart template to a chart in Aspose.Cells C# | Aspose.Cells change chart template from byte array example | set custom font color for chart data labels using Aspose.Cells .NET | add chart with template overload Aspose.Cells C# sample code | modify data label position to InsideEnd in Aspose.Cells chart
// Tags: Aspose.Cells ChangeTemplate chart API | Aspose.Cells Add chart from byte array | custom data label font color Aspose.Cells | chart data label position InsideEnd Aspose.Cells | load .crtx template Aspose.Cells .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartTemplateDemo
{
    // The example loads a .crtx chart template into a byte array, creates two workbooks, and demonstrates two ways to apply the template to column charts: using ChangeTemplate and using the Add(byte[], ...) overload. It also customizes data label settings—showing values, setting a dark‑blue font, and positioning labels inside the end—before saving the workbooks as .xlsx files.
    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Load chart template bytes if the file exists.
            // ------------------------------------------------------------
            byte[] templateBytes = null;
            string templatePath = "ChartTemplate.crtx";

            try
            {
                if (File.Exists(templatePath))
                {
                    templateBytes = File.ReadAllBytes(templatePath);
                }
                else
                {
                    Console.WriteLine($"Template file '{templatePath}' not found. Continuing without template.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading template file: {ex.Message}");
                // Continue without template
            }

            // ------------------------------------------------------------
            // 2. Create a workbook and add a chart that will use the template
            //    via the ChangeTemplate method (if template is available).
            // ------------------------------------------------------------
            try
            {
                Workbook wbChangeTemplate = new Workbook();
                Worksheet wsChange = wbChangeTemplate.Worksheets[0];

                // Sample data
                wsChange.Cells["A1"].PutValue("Category");
                wsChange.Cells["A2"].PutValue("A");
                wsChange.Cells["A3"].PutValue("B");
                wsChange.Cells["A4"].PutValue("C");
                wsChange.Cells["B1"].PutValue("Value");
                wsChange.Cells["B2"].PutValue(10);
                wsChange.Cells["B3"].PutValue(20);
                wsChange.Cells["B4"].PutValue(30);

                // Add a basic column chart
                int chartIdx1 = wsChange.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart1 = wsChange.Charts[chartIdx1];
                chart1.SetChartDataRange("A1:B4", true);

                // Apply the template if we have it
                if (templateBytes != null)
                {
                    chart1.ChangeTemplate(templateBytes);
                }

                // Adjust data label properties
                Series series1 = chart1.NSeries[0];
                series1.DataLabels.ShowValue = true;               // ensure values are shown
                series1.DataLabels.Font.Color = Color.DarkBlue;    // override font color
                series1.DataLabels.ApplyFont();                    // apply font to all labels

                // Save the workbook
                wbChangeTemplate.Save("Chart_With_ChangedTemplate.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating workbook with ChangeTemplate: {ex.Message}");
            }

            // ------------------------------------------------------------
            // 3. Create another workbook and add a chart directly with the
            //    template using the Add(byte[], ...) overload (if template is available).
            // ------------------------------------------------------------
            try
            {
                Workbook wbAddTemplate = new Workbook();
                Worksheet wsAdd = wbAddTemplate.Worksheets[0];

                // Sample data (same structure)
                wsAdd.Cells["A1"].PutValue("Category");
                wsAdd.Cells["A2"].PutValue("A");
                wsAdd.Cells["A3"].PutValue("B");
                wsAdd.Cells["A4"].PutValue("C");
                wsAdd.Cells["B1"].PutValue("Value");
                wsAdd.Cells["B2"].PutValue(15);
                wsAdd.Cells["B3"].PutValue(25);
                wsAdd.Cells["B4"].PutValue(35);

                int chartIdx2;

                if (templateBytes != null)
                {
                    // Add a chart using the template byte array.
                    // Parameters: template data, data range, isVertical, topRow, leftColumn, bottomRow, rightColumn
                    chartIdx2 = wsAdd.Charts.Add(
                        templateBytes,   // template byte array
                        "A1:B4",         // data range
                        true,            // plot by column
                        5, 0, 20, 8);    // position of the chart
                }
                else
                {
                    // Fallback: add a regular chart without a template.
                    chartIdx2 = wsAdd.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                    wsAdd.Charts[chartIdx2].SetChartDataRange("A1:B4", true);
                }

                Chart chart2 = wsAdd.Charts[chartIdx2];

                // Optionally modify a specific setting.
                chart2.NSeries[0].DataLabels.Position = LabelPositionType.InsideEnd;

                // Save the workbook
                wbAddTemplate.Save("Chart_Added_With_Template.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating workbook with Add template overload: {ex.Message}");
            }
        }
    }
}
