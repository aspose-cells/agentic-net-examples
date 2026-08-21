// Title: C# Example: Convert Excel to HTML with Aspose.Cells – Preserve Conditional Formatting (Default HtmlSaveOptions)
// Description: A complete C# snippet that checks an XLSX file, loads it with Aspose.Cells, optionally adds a conditional‑format rule, and saves the workbook as HTML using the default HtmlSaveOptions, which automatically export conditional formatting.
// Keywords: Aspose.Cells | C# Excel to HTML | conditional formatting export | HtmlSaveOptions default | convert .xlsx to .html | Aspose.Cells sample code | .NET workbook conversion | HTML report from Excel | GitHub Aspose.Cells example | code snippet
// Common Searches: Aspose.Cells export conditional formatting to HTML | C# default HtmlSaveOptions Excel to HTML conversion | How to keep conditional formatting when saving Excel as HTML | Aspose.Cells HTML conversion example .NET | Convert .xlsx to .html preserving styles C#
// Developer Intent: Generate an HTML file from an Excel workbook in C# while retaining any conditional formatting using Aspose.Cells.
// Use Cases: Display a spreadsheet on a web page with its color‑coded rules intact. | Create automated HTML reports from Excel templates that rely on conditional formatting. | Batch‑process multiple .xlsx files into web‑ready HTML without losing formatting.
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells and saves it as HTML using the default HtmlSaveOptions, ensuring conditional formatting is included. | Explain how Aspose.Cells' HtmlSaveOptions handles conditional formatting during Excel‑to‑HTML conversion. | Show how to modify the example to export each worksheet to a separate HTML file while preserving conditional formatting.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

// A complete C# snippet that checks an XLSX file, loads it with Aspose.Cells, optionally adds a conditional‑format rule, and saves the workbook as HTML using the default HtmlSaveOptions, which automatically export conditional formatting.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // ------------------------------------------------------------
            // OPTIONAL: Add a sample conditional formatting rule so that the
            // conversion demonstrates exporting conditional formatting.
            // This step can be omitted if the source file already contains
            // conditional formatting.
            // ------------------------------------------------------------
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range to which the conditional formatting will be applied
            CellArea area = CellArea.CreateCellArea("A1", "B10");

            // Add a new ConditionalFormatting collection for the defined range
            int cfIndex = sheet.ConditionalFormattings.Add();
            var cf = sheet.ConditionalFormattings[cfIndex];
            cf.AddArea(area);

            // Create a condition: cells with values > 50 will have a red background
            int conditionIndex = cf.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "50", null);
            var condition = cf[conditionIndex];
            var style = condition.Style;
            style.ForegroundColor = Color.Red;
            style.Pattern = BackgroundType.Solid;
            condition.Style = style;

            // ------------------------------------------------------------
            // Save the workbook as HTML using default HtmlSaveOptions.
            // The default options include exporting conditional formatting.
            // ------------------------------------------------------------
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(); // default options
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook successfully saved as HTML to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
