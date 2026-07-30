// Title: Apply a .crtx Chart Template and Customize Data Labels in Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add sample data, insert a column chart, load a .crtx chart template (if available), apply the template, enable data labels, set their font color and size, and save the file as XLSX using Aspose.Cells for C#.
// Keywords: Aspose.Cells | .crtx chart template | chart template C# | data labels formatting | column chart Aspose.Cells | ChangeTemplate method | SetChartDataRange | Series DataLabels | Excel automation .NET | chart styling programmatically
// Common Searches: how to load a .crtx chart template with Aspose.Cells | set data label font color and size in Aspose.Cells chart | create column chart from worksheet data C# Aspose.Cells | apply predefined chart template to multiple charts Aspose.Cells | Aspose.Cells change chart template example
// Developer Intent: Load a chart template and programmatically format data labels for a newly created chart in Aspose.Cells.
// Use Cases: Standardize chart appearance across generated reports by reusing a .crtx template. | Automatically enable and style data labels for series in column charts. | Produce Excel workbooks where charts adopt predefined formatting without manual editing.
// AI Prompts: Generate C# code that creates a line chart, applies a .crtx template, and sets data labels to show percentages in red using Aspose.Cells. | Show how to apply the same chart template to several charts in a workbook while customizing each chart's data label size. | Explain how to build a chart template programmatically with Aspose.Cells and use it without an external .crtx file.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsChartTemplateDemo
{
    // Demonstrates how to create a workbook, add sample data, insert a column chart, load a .crtx chart template (if available), apply the template, enable data labels, set their font color and size, and save the file as XLSX using Aspose.Cells for C#.
    public class Program
    {
        public static void Main()
        {
            try
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
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.SetChartDataRange("A1:B4", true);

                // Load a pre‑created chart template (.crtx) if it exists
                const string templatePath = "ChartTemplate.crtx";
                if (File.Exists(templatePath))
                {
                    byte[] templateData = File.ReadAllBytes(templatePath);
                    chart.ChangeTemplate(templateData);
                }
                else
                {
                    Console.WriteLine($"Template file '{templatePath}' not found. Continuing without applying a template.");
                }

                // Ensure data labels are visible and customize their appearance
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;               // Show the value in each label
                series.DataLabels.Font.Color = Color.DarkBlue;    // Set font color
                series.DataLabels.Font.Size = 12;                 // Set font size
                series.DataLabels.ApplyFont();                    // Apply the font settings

                // Save the workbook with the chart
                workbook.Save("ChartWithTemplate.xlsx", SaveFormat.Xlsx);
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
