// Title: Apply a greater‑than conditional formatting rule to a cell range and export the worksheet as HTML with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add conditional formatting before converting workbook to HTML | preserve conditional formatting colors when saving Excel as HTML with Aspose.Cells | how to export only the active worksheet to HTML using HtmlSaveOptions in .NET | embed worksheet images as Base64 in HTML output with Aspose.Cells | apply greater than value conditional formatting to range A1:A10 in Aspose.Cells
// Tags: conditional-formatting export html Aspose.Cells | htmlsaveoptions exportactiveworksheetonly c# | exportimagesasbase64 aspose.cells html | formatcondition cellvalue greaterthan c# | style foregroundcolor lightsalmon aspose

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads or creates an Excel workbook, defines a conditional formatting rule that highlights cells with values greater than 100 using a LightSalmon background on range A1:A10, configures HtmlSaveOptions to embed images as Base64 and export only the active worksheet, and saves the workbook as an HTML file while preserving the conditional formatting.
class ConditionalFormattingToHtml
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.html";

            // Load existing workbook or create a simple one if it does not exist
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                for (int i = 0; i < 10; i++)
                {
                    ws.Cells[i, 0].PutValue(i * 20); // 0,20,...180
                }
                workbook.Save(inputPath);
            }

            // First worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define range A1:A10 for conditional formatting
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 9,
                EndColumn = 0
            };

            // Add a conditional formatting rule: value > 100 → LightSalmon background
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection conditions = sheet.ConditionalFormattings[cfIndex];
            conditions.AddArea(area);
            FormatCondition condition = conditions[conditions.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.GreaterThan,
                "100",
                null)]; // formula2 not required

            // Style for the condition
            Style style = workbook.CreateStyle();
            style.ForegroundColor = Color.LightSalmon;
            style.Pattern = BackgroundType.Solid;
            condition.Style = style;

            // HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportImagesAsBase64 = true,
                ExportActiveWorksheetOnly = true
                // Conditional formatting is exported by default
            };

            // Save as HTML
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
