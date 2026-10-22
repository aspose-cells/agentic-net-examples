// Title: Export an Excel worksheet to HTML with gridlines and conditional formatting using Aspose.Cells for .NET
// AI Prompts: Generate C# code that adds conditional formatting to a range and saves the workbook as HTML with gridlines enabled via Aspose.Cells. | Demonstrate how to configure HtmlSaveOptions to include gridlines and embed images as Base64 while preserving conditional formatting in the HTML output.
// Common Searches: how to export Excel to HTML with gridlines using Aspose.Cells C# | Aspose.Cells conditional formatting retained in HTML export | HtmlSaveOptions ExportGridLines true example for .NET | embed images as Base64 in Aspose.Cells HTML output | save only active worksheet as HTML with gridlines Aspose.Cells
// Tags: Aspose.Cells gridlines HTML export | conditional formatting HTML export .NET | HtmlSaveOptions gridlines option C# | embed images Base64 Aspose.Cells | export active worksheet only HTML Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsExportExample
{
    // The program creates a workbook, adds sample scores, applies two conditional formatting rules (green for scores ≥80 and coral for scores <60) to the range A2:A5, configures HtmlSaveOptions to export gridlines and embed images as Base64, and saves the active worksheet as an HTML file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Data";

                // Populate sample data
                sheet.Cells["A1"].PutValue("Score");
                sheet.Cells["A2"].PutValue(85);
                sheet.Cells["A3"].PutValue(70);
                sheet.Cells["A4"].PutValue(55);
                sheet.Cells["A5"].PutValue(40);

                // Define the range for conditional formatting (A2:A5)
                CellArea area = new CellArea
                {
                    StartRow = 1,   // Row 2 (zero‑based)
                    EndRow = 4,     // Row 5
                    StartColumn = 0,
                    EndColumn = 0
                };

                // Add a new ConditionalFormatting object to the worksheet
                int cfIndex = sheet.ConditionalFormattings.Add();
                var cf = sheet.ConditionalFormattings[cfIndex];
                cf.AddArea(area);

                // Condition: scores >= 80 → light green background
                // Using GreaterThan with 79 to emulate >=80 (compatible with older API versions)
                int highCondIdx = cf.AddCondition(
                    FormatConditionType.CellValue,
                    OperatorType.GreaterThan,
                    "79",
                    null);
                FormatCondition conditionHigh = cf[highCondIdx];
                conditionHigh.Style.ForegroundColor = Color.LightGreen;
                conditionHigh.Style.Pattern = BackgroundType.Solid;

                // Condition: scores < 60 → light coral background
                int lowCondIdx = cf.AddCondition(
                    FormatConditionType.CellValue,
                    OperatorType.LessThan,
                    "60",
                    null);
                FormatCondition conditionLow = cf[lowCondIdx];
                conditionLow.Style.ForegroundColor = Color.LightCoral;
                conditionLow.Style.Pattern = BackgroundType.Solid;

                // Set HTML save options (grid lines and images as Base64)
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    ExportGridLines = true,
                    ExportImagesAsBase64 = true,
                    ExportActiveWorksheetOnly = true
                };

                // Save the workbook as HTML
                string outputPath = "ExportedWithGridLinesAndFormatting.html";
                workbook.Save(outputPath, htmlOptions);

                Console.WriteLine($"Workbook exported to HTML at: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
