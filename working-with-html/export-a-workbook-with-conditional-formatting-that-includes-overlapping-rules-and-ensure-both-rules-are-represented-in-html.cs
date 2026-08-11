// Title: Export Workbook with Overlapping Conditional Formatting to HTML – Aspose.Cells for .NET
// Description: Creates a workbook, fills column A with numbers, adds two overlapping conditional‑formatting rules (yellow for 15‑45, orange for >30), configures HtmlSaveOptions to keep all styles and export CSS separately, and saves the result as an HTML file that shows both rules.
// Keywords: Aspose.Cells | C# | conditional formatting | HTML export | overlapping rules | HtmlSaveOptions | ExportWorksheetCSSSeparately | ExcludeUnusedStyles | Excel to HTML | style preservation
// Common Searches: Aspose.Cells export overlapping conditional formatting to HTML | preserve multiple conditional formats in HTML output .NET | HtmlSaveOptions ExportWorksheetCSSSeparately example | how to keep all conditional formatting rules when saving as HTML | C# Aspose.Cells conditional formatting HTML export
// Developer Intent: Generate an HTML file from a workbook while ensuring that every overlapping conditional‑formatting rule is retained in the output.
// Use Cases: Produce an HTML report that visually reflects two overlapping formatting rules on numeric data. | Separate conditional‑formatting CSS into its own file for easier inspection or customization. | Debug or document Excel styling by preserving all conditional formats in the HTML conversion.
// AI Prompts: Write C# code with Aspose.Cells that exports a workbook containing overlapping conditional‑formatting rules to HTML, showing both styles. | Explain the impact of HtmlSaveOptions properties ExportWorksheetCSSSeparately and ExcludeUnusedStyles on overlapping conditional formats in the generated HTML. | Demonstrate how to adjust the StopIfTrue property to control rule precedence when exporting conditional formatting with Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsConditionalFormattingHtml
{
    // Creates a workbook, fills column A with numbers, adds two overlapping conditional‑formatting rules (yellow for 15‑45, orange for >30), configures HtmlSaveOptions to keep all styles and export CSS separately, and saves the result as an HTML file that shows both rules.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample numeric data in column A (A1:A10)
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i, 0].PutValue(i * 10); // 0,10,20,...,90
            }

            // Add a conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection cfCollection = sheet.ConditionalFormattings[cfIndex];

            // Define the range that both rules will apply to (A1:A10)
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 9,
                StartColumn = 0,
                EndColumn = 0
            };
            cfCollection.AddArea(area);

            // First rule: values between 15 and 45 -> yellow background
            int condIdx1 = cfCollection.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.Between,
                "15",
                "45");
            FormatCondition cond1 = cfCollection[condIdx1];
            cond1.Style.BackgroundColor = Color.Yellow;
            cond1.StopIfTrue = false; // allow lower‑priority rules to be evaluated

            // Second rule: values greater than 30 -> orange background
            // This overlaps with the first rule for values 31‑45
            int condIdx2 = cfCollection.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.GreaterThan,
                "30",
                null);
            FormatCondition cond2 = cfCollection[condIdx2];
            cond2.Style.BackgroundColor = Color.Orange;
            cond2.StopIfTrue = false;

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Export CSS separately so that conditional formatting styles are clearly visible
                ExportWorksheetCSSSeparately = true,
                // Keep all generated styles (including those from overlapping rules)
                ExcludeUnusedStyles = false
            };

            // Save the workbook as HTML
            string outputPath = "ConditionalFormattingOverlap.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook saved to HTML at: {outputPath}");
        }
    }
}
