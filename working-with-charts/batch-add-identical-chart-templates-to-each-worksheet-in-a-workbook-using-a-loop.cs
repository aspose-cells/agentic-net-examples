// Title: Add a .crtx Chart Template to Every Worksheet in an Aspose.Cells Workbook Using C# Loop
// Description: C# example that loads a .crtx chart template (if available), creates a workbook with three sheets of sample data, loops through each sheet to add the template chart (or a default column chart as fallback), sets a dynamic title, and saves the file as XLSX.
// Keywords: Aspose.Cells | C# | chart template | .crtx | batch chart creation | multiple worksheets | loop | add chart | default chart fallback | Excel XLSX | sample data | dynamic chart title
// Common Searches: Aspose.Cells apply same chart template to all sheets C# | C# loop add chart to each worksheet Aspose.Cells | use .crtx template with Aspose.Cells | batch create charts in Excel workbook C# | fallback to default chart when template missing Aspose.Cells
// Developer Intent: Programmatically apply one chart template to every worksheet in a workbook.
// Use Cases: Create a quarterly sales report where each sheet shows a column chart styled by a shared .crtx template. | Build a multi‑sheet dashboard with consistent chart formatting across all tabs. | Generate workbooks that automatically use a default chart if the specified template file cannot be found.
// AI Prompts: Generate C# code that loads a .crtx file and adds the same chart to all worksheets in an Aspose.Cells workbook. | Explain how to handle a missing chart template gracefully while looping through worksheets in Aspose.Cells. | Show how to assign a chart title that includes the worksheet name for each chart added in a batch operation.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartTemplateBatch
{
    // C# example that loads a .crtx chart template (if available), creates a workbook with three sheets of sample data, loops through each sheet to add the template chart (or a default column chart as fallback), sets a dynamic title, and saves the file as XLSX.
    class Program
    {
        static void Main()
        {
            try
            {
                // Load chart template if it exists
                byte[] templateData = null;
                const string templatePath = "ChartTemplate.crtx";
                if (File.Exists(templatePath))
                {
                    templateData = File.ReadAllBytes(templatePath);
                }
                else
                {
                    Console.WriteLine($"Warning: Template file '{templatePath}' not found. Charts will be created without a template.");
                }

                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add three worksheets with sample data
                for (int i = 0; i < 3; i++)
                {
                    Worksheet sheet;
                    if (i < workbook.Worksheets.Count)
                    {
                        sheet = workbook.Worksheets[i];
                        sheet.Name = $"Sheet{i + 1}";
                    }
                    else
                    {
                        sheet = workbook.Worksheets.Add($"Sheet{i + 1}");
                    }

                    sheet.Cells["A1"].PutValue("Category");
                    sheet.Cells["B1"].PutValue("Value");
                    for (int row = 2; row <= 5; row++)
                    {
                        sheet.Cells[$"A{row}"].PutValue($"Item {row - 1}");
                        sheet.Cells[$"B{row}"].PutValue(row * 10);
                    }
                }

                // Add a chart to each worksheet
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    int chartIndex;
                    if (templateData != null)
                    {
                        // Add chart using the template
                        chartIndex = ws.Charts.Add(
                            templateData,   // chart template data
                            "A1:B5",        // data range
                            true,           // plot series by column
                            5, 0, 20, 8);   // position (topRow, leftColumn, bottomRow, rightColumn)
                    }
                    else
                    {
                        // Add a default column chart when template is unavailable
                        chartIndex = ws.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                        ws.Charts[chartIndex].NSeries.Add("=Sheet!$B$2:$B$5", true);
                        ws.Charts[chartIndex].NSeries.CategoryData = "=Sheet!$A$2:$A$5";
                    }

                    // Customize chart title
                    Chart chart = ws.Charts[chartIndex];
                    chart.Title.Text = $"Sample Chart on {ws.Name}";
                }

                // Save the workbook
                const string outputPath = "WorkbookWithCharts.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
