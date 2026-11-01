// Title: Convert an Excel workbook to self‑contained HTML while preserving conditional formatting colors using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export Excel to HTML with conditional formatting colors retained | How to keep conditional format font color when saving workbook as HTML in .NET | C# example for converting .xlsx to HTML with conditional formatting using Aspose.Cells | HtmlSaveOptions preserve cell style colors Aspose.Cells C# | Render Excel conditional formatting in HTML output with Aspose.Cells
// Tags: Aspose.Cells HTML export preserving conditional formatting | C# workbook to HTML conversion with style retention | HtmlSaveOptions self-contained HTML output | conditional formatting color support in Aspose.Cells HTML | Excel to HTML conversion using Aspose.Cells .NET

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

// The sample creates or loads an Excel file, adds a conditional formatting rule that colors the font red for values greater than 200, and saves the workbook as a self‑contained HTML document using Aspose.Cells HtmlSaveOptions with ExportImagesAsBase64 enabled, ensuring the conditional formatting colors appear correctly in the HTML.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Ensure the input file exists; create a simple workbook if it does not.
            if (!File.Exists(inputPath))
            {
                var wb = new Workbook();
                var ws = wb.Worksheets[0];
                ws.Cells["A1"].PutValue("Sample Data");
                ws.Cells["A2"].PutValue(123);
                ws.Cells["A3"].PutValue(456);

                // Add a simple conditional format for demonstration.
                int cfIndex = ws.ConditionalFormattings.Add();
                var cf = ws.ConditionalFormattings[cfIndex];
                // Add a condition: cell value > 200
                int conditionIndex = cf.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "200", null);
                var condition = cf[conditionIndex];
                condition.Style.Font.Color = Color.Red;

                wb.Save(inputPath);
            }

            // Load the source workbook.
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Embed images as Base64 to keep the HTML self‑contained.
                ExportImagesAsBase64 = true
            };

            // Save the workbook as HTML.
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully converted to HTML: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
