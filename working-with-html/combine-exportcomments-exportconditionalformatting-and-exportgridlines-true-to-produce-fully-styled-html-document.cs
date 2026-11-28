// Title: Generate a fully styled HTML file from an Excel workbook with comments, conditional formatting, and grid lines using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export workbook to HTML with comments and conditional formatting | How to include grid lines when saving Excel as HTML using Aspose.Cells .NET | C# preserve cell comments and conditional formatting in HTML output with Aspose.Cells | Export only the active worksheet to HTML with embedded images using Aspose.Cells | Configure HtmlSaveOptions to embed images as Base64 in HTML export
// Tags: htmlsaveoptions exportgridlines setting | export comments to html using Aspose.Cells | conditional formatting preservation in html export | embed images as base64 in html output Aspose.Cells | active worksheet only html export Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

// The example creates a workbook, fills sample data, adds a comment to B2, applies yellow conditional formatting to B1:B10 for values greater than 50, and uses HtmlSaveOptions (ExportGridLines, ExportActiveWorksheetOnly, ExportImagesAsBase64, PageTitle) to save the workbook as a fully styled HTML file named StyledExport.html.
class Program
{
    static void Main()
    {
        try
        {
            // -----------------------------------------------------------------
            // 1. Create a workbook and fill it with sample data
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate 10 rows of data in columns A and B
            for (int row = 0; row < 10; row++)
            {
                cells[row, 0].PutValue(row + 1);               // Column A: 1..10
                cells[row, 1].PutValue((row + 1) * 10);        // Column B: 10,20,...,100
            }

            // -----------------------------------------------------------------
            // 2. Add a comment to demonstrate ExportComments (enabled by default)
            // -----------------------------------------------------------------
            int commentIndex = sheet.Comments.Add("B2");
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = "This is a sample comment that will be exported to HTML.";

            // -----------------------------------------------------------------
            // 3. Add conditional formatting to demonstrate ExportConditionalFormatting (enabled by default)
            // -----------------------------------------------------------------
            int cfIndex = sheet.ConditionalFormattings.Add();
            var cf = sheet.ConditionalFormattings[cfIndex];

            // Define the range B1:B10 (zero‑based indices)
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 1,
                EndRow = 9,
                EndColumn = 1
            };
            cf.AddArea(area);

            // Add the condition (CellValue > 50)
            int conditionIdx = cf.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.GreaterThan,
                "50",
                null);
            var condition = cf[conditionIdx];
            condition.Style.BackgroundColor = Color.Yellow;

            // -----------------------------------------------------------------
            // 4. Configure HTML save options to export grid lines (comments and conditional formatting are exported by default)
            // -----------------------------------------------------------------
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportGridLines = true,               // Show grid lines
                ExportActiveWorksheetOnly = true,     // Export only the active sheet
                ExportImagesAsBase64 = true,          // Embed images directly
                PageTitle = "Fully Styled HTML Export"
            };

            // -----------------------------------------------------------------
            // 5. Save the workbook as a fully styled HTML document
            // -----------------------------------------------------------------
            string outputHtml = "StyledExport.html";
            workbook.Save(outputHtml, htmlOptions);

            Console.WriteLine($"HTML file successfully saved to: {outputHtml}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred:");
            Console.WriteLine(ex.Message);
        }
    }
}
