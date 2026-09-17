// Title: Save an Excel workbook as HTML while preserving conditional formatting background colors using Aspose.Cells for .NET
// AI Prompts: Write C# code that adds a conditional formatting rule with a solid LightGreen fill and then saves the workbook to an HTML file, ensuring the formatting appears in the output. | Show how to configure Aspose.Cells HtmlSaveOptions so that conditional formatting styles are included when exporting a worksheet to HTML.
// Common Searches: Aspose.Cells .NET keep conditional formatting colors when converting Excel to HTML | C# export Excel to HTML with cell background colors from conditional rules | HtmlSaveOptions setting to retain conditional formatting in generated HTML | How to preserve conditional formatting fill colors in HTML output using Aspose.Cells
// Tags: export workbook to HTML with conditional formatting | Aspose.Cells HtmlSaveOptions preserve cell fill | C# conditional formatting background export | save Excel as HTML retaining styles | conditional formatting to HTML Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a workbook, applies a conditional formatting rule that fills cells with LightGreen when the value exceeds 15, and saves the workbook as an HTML file using HtmlSaveOptions so that the conditional formatting background colors are retained in the generated HTML.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);

            // Define the range for conditional formatting (A1:A3)
            CellArea range = CellArea.CreateCellArea("A1", "A3");

            // Add a new conditional formatting rule to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            var cf = sheet.ConditionalFormattings[cfIndex];
            cf.AddArea(range);

            // Add a condition: if cell value > 15, apply a style
            int conditionIndex = cf.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.GreaterThan,
                "15",
                string.Empty); // second formula not required

            // Create a style with a background color
            Style bgStyle = workbook.CreateStyle();
            bgStyle.ForegroundColor = Color.LightGreen;
            bgStyle.Pattern = BackgroundType.Solid;

            // Assign the style to the condition
            cf[conditionIndex].Style = bgStyle;

            // Configure HTML save options (conditional formatting is exported by default)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Ensure the output directory exists
            string outputPath = "output.html";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as HTML, preserving background colors from conditional formatting
            workbook.Save(outputPath, htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
