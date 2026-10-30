// Title: Export an Excel workbook with overlapping conditional formatting rules to HTML using Aspose.Cells for .NET
// Common Searches: how to export Excel with multiple conditional formatting rules to HTML using Aspose.Cells C# | Aspose.Cells preserve overlapping conditional formats when saving as HTML | C# save workbook as HTML while keeping both >50 and <20 conditional styles | export range A1:A10 with two conditional formatting conditions to HTML with Aspose.Cells | HtmlSaveOptions retain all conditional formatting rules in generated HTML file
// Tags: conditional formatting overlapping HTML export | Aspose.Cells HTMLSaveOptions preserve conditional formats | C# export Excel with multiple format conditions | Excel to HTML conversion retaining conditional styling | save workbook as HTML with dual conditional rules

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, fills cells A1:A10 with numeric values, applies two overlapping conditional formatting rules (values > 50 → red background/white font, values < 20 → green background/black font) on the same range, and saves the workbook as an HTML file using Aspose.Cells' HtmlSaveOptions, ensuring both rules appear in the generated HTML.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data in column A (values 0,10,...,90)
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i, 0].PutValue(i * 10);
            }

            // Define the cell area (A1:A10) for conditional formatting
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 9,
                StartColumn = 0,
                EndColumn = 0
            };

            // ---------- First conditional formatting rule ----------
            // Values greater than 50 -> red background, white font
            int cfIndex1 = sheet.ConditionalFormattings.Add();
            dynamic cf1 = sheet.ConditionalFormattings[cfIndex1];
            cf1.AddArea(area);
            int conditionIndex1 = cf1.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "50");
            dynamic condition1 = cf1[conditionIndex1];
            condition1.Style.BackgroundColor = Color.Red;
            condition1.Style.Font.Color = Color.White;

            // ---------- Second conditional formatting rule ----------
            // Values less than 20 -> green background, black font
            int cfIndex2 = sheet.ConditionalFormattings.Add();
            dynamic cf2 = sheet.ConditionalFormattings[cfIndex2];
            cf2.AddArea(area);
            int conditionIndex2 = cf2.AddCondition(FormatConditionType.CellValue, OperatorType.LessThan, "20");
            dynamic condition2 = cf2[conditionIndex2];
            condition2.Style.BackgroundColor = Color.Green;
            condition2.Style.Font.Color = Color.Black;

            // Save the workbook as HTML (conditional formatting is exported by default)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            string outputPath = "ConditionalFormatting.html";
            // Ensure the directory exists (if a path is provided)
            string directory = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
