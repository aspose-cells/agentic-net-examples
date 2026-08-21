// Title: C# – Export Excel to HTML with Conditional Formatting Preserved using Aspose.Cells
// Description: Shows how to build a workbook, apply a conditional formatting rule (values > 5 highlighted in light‑green), set HtmlSaveOptions to retain merged cells, grid lines and styles, and save the result as HTML while keeping the visual formatting intact with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | conditional formatting | HTML export | HtmlSaveOptions | preserve cell styles | Excel to HTML | grid lines | merged cells | style preservation | programmatic Excel | Aspose.Cells example
// Common Searches: Aspose.Cells keep conditional formatting in HTML | Export Excel with colors to HTML .NET | HtmlSaveOptions preserve cell styles | C# conditional formatting to HTML | retain merged cells when converting Excel to HTML | Aspose.Cells HTML export example
// Developer Intent: Add a conditional formatting rule to a worksheet and export it as HTML while maintaining the visual appearance.
// Use Cases: Generate a web‑ready report where values above a threshold are highlighted. | Publish Excel data to a portal with grid lines and merged‑cell layout intact. | Automate creation of HTML dashboards that reflect Excel conditional styles.
// AI Prompts: Provide a C# snippet that adds multiple conditional formatting rules and exports the workbook to HTML with Aspose.Cells. | Explain how to embed CSS in the HTML output while preserving conditional formatting using HtmlSaveOptions. | Show how to load an existing .xlsx file, apply a new conditional format, and save it as HTML with visual styles retained.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsConditionalFormattingToHtml
{
    // Shows how to build a workbook, apply a conditional formatting rule (values > 5 highlighted in light‑green), set HtmlSaveOptions to retain merged cells, grid lines and styles, and save the result as HTML while keeping the visual formatting intact with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (values 1 to 10) in column A
                for (int i = 0; i < 10; i++)
                {
                    sheet.Cells[i, 0].PutValue(i + 1);
                }

                // Add a conditional formatting rule for the range A1:A10
                int cfIndex = sheet.ConditionalFormattings.Add(); // create a new CF entry
                var cf = sheet.ConditionalFormattings[cfIndex];

                // Define the target range using CellArea (A1:A10 => rows 0‑9, column 0)
                CellArea area = new CellArea
                {
                    StartRow = 0,
                    StartColumn = 0,
                    EndRow = 9,
                    EndColumn = 0
                };
                cf.AddArea(area);

                // Condition: cell value greater than 5
                int conditionIndex = cf.AddCondition(
                    FormatConditionType.CellValue,
                    OperatorType.GreaterThan,
                    "5",
                    string.Empty); // second formula not required for this operator

                // Define the style to apply (light green background)
                Style style = workbook.CreateStyle();
                style.ForegroundColor = Color.LightGreen;
                style.Pattern = BackgroundType.Solid;
                cf[conditionIndex].Style = style;

                // Configure HTML save options to preserve visual formatting
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    MergeAreas = true,                     // Merge merged cells
                    ExportActiveWorksheetOnly = false,    // Export whole workbook
                    ExportGridLines = true,                // Include grid lines
                    DisableCss = false,                    // Keep external CSS
                    PresentationPreference = true,        // Presentation‑friendly output
                    PageTitle = "Conditional Formatting Demo"
                };

                // Determine output path and ensure directory exists
                string outputPath = "ConditionalFormatting.html";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as HTML
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook saved to '{outputPath}' with conditional formatting preserved.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
