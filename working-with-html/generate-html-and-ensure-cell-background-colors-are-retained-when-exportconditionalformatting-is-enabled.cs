// Title: How to export Excel conditional formatting with background colors to HTML using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export conditional formatting to HTML preserving cell background colors | Enable ExportConditionalFormatting in HtmlSaveOptions to keep red background in HTML output | C# generate HTML from workbook with conditional formatting colors using Aspose.Cells | Retain conditional formatting colors when saving Excel as HTML with Aspose.Cells .NET
// Tags: Aspose.Cells HtmlSaveOptions ExportConditionalFormatting | Aspose.Cells conditional formatting HTML export | C# export Excel to HTML with background colors | Aspose.Cells preserve conditional formatting colors | HtmlSaveOptions ExportImagesAsBase64

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsHtmlExportDemo
{
    // The example creates a workbook, adds a conditional formatting rule that applies a solid red background to cells with values greater than 15, configures HtmlSaveOptions to enable ExportConditionalFormatting, exports images as Base64, disables separate worksheet CSS, and saves the workbook as an HTML file while preserving the conditional formatting background colors.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Output HTML file path
                string outputHtmlPath = "ConditionalFormattingDemo.html";

                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Item");
                sheet.Cells["B1"].PutValue("Quantity");
                sheet.Cells["A2"].PutValue("Apples");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("Bananas");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("Cherries");
                sheet.Cells["B4"].PutValue(30);

                // ------------------------------------------------------------
                // Add conditional formatting (red background for values > 15)
                // ------------------------------------------------------------
                int cfIndex = sheet.ConditionalFormattings.Add();

                // The collection returns a FormatConditionCollection for each conditional formatting rule
                FormatConditionCollection cf = sheet.ConditionalFormattings[cfIndex];

                // Define the range B2:B4
                CellArea area = new CellArea
                {
                    StartRow = 1,   // B2 (zero‑based)
                    StartColumn = 1,
                    EndRow = 3,     // B4
                    EndColumn = 1
                };
                cf.AddArea(area);

                // Add the condition (cell value > 15)
                int conditionIdx = cf.AddCondition(
                    FormatConditionType.CellValue,
                    OperatorType.GreaterThan,
                    "15",
                    null);

                // Set the style (red background)
                FormatCondition condition = cf[conditionIdx];
                Style style = workbook.CreateStyle();
                style.ForegroundColor = Color.Red;
                style.Pattern = BackgroundType.Solid;
                condition.Style = style;

                // ------------------------------------------------------------
                // Prepare HTML save options
                // ------------------------------------------------------------
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

                // Export conditional formatting if the property exists
                var exportCfProp = typeof(HtmlSaveOptions).GetProperty("ExportConditionalFormatting");
                if (exportCfProp != null && exportCfProp.CanWrite)
                {
                    exportCfProp.SetValue(htmlOptions, true);
                }

                // Export images as Base64 strings
                htmlOptions.ExportImagesAsBase64 = true;

                // Ensure generated images are not transparent
                htmlOptions.ImageOptions.Transparent = false;

                // Keep CSS inline (do not separate per worksheet) if the property exists
                var cssProp = typeof(HtmlSaveOptions).GetProperty("ExportWorksheetCSSSeparately");
                if (cssProp != null && cssProp.CanWrite)
                {
                    cssProp.SetValue(htmlOptions, false);
                }

                // ------------------------------------------------------------
                // Save the workbook as HTML
                // ------------------------------------------------------------
                workbook.Save(outputHtmlPath, htmlOptions);

                Console.WriteLine($"HTML file generated at: {Path.GetFullPath(outputHtmlPath)}");
                Console.WriteLine("Conditional formatting with background colors has been exported.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
