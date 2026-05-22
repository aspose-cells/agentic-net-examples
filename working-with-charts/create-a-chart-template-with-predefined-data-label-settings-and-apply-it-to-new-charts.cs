using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartTemplateDemo
{
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

                // Add a basic column chart (the template will later change its appearance)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];
                chart.SetChartDataRange("A1:B4", true);

                // Load a pre‑created chart template (.crtx) if it exists
                const string templatePath = "ChartTemplate.crtx";
                if (File.Exists(templatePath))
                {
                    try
                    {
                        // Aspose.Cells expects the template as a byte array
                        byte[] templateData = File.ReadAllBytes(templatePath);
                        chart.ChangeTemplate(templateData);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to apply template: {ex.Message}");
                    }
                }
                else
                {
                    Console.WriteLine($"Template file '{templatePath}' not found. Continuing without applying a template.");
                }

                // Modify individual data‑label properties if needed
                Series series = chart.NSeries[0];
                series.DataLabels.Font.Color = Color.DarkBlue;
                series.DataLabels.Font.Size = 12;
                series.DataLabels.ApplyFont(); // Apply the font settings to all child label nodes

                // Save the workbook with the chart that now uses the template settings (if applied)
                const string outputPath = "ChartWithTemplateApplied.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}