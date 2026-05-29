using System;
using Aspose.Cells;

namespace AsposeCellsHtmlValidation
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data
            sheet.Cells["A1"].PutValue("Hello World");               // plain text
            sheet.Cells["B1"].PutValue(12345);                       // numeric value
            sheet.Cells["C1"].Formula = "=SUM(B1,B1)";               // formula

            // Apply some formatting (bold and red font) to demonstrate CSS/inline styles
            Style style = sheet.Cells["A1"].GetStyle();
            style.Font.IsBold = true;
            style.Font.Color = System.Drawing.Color.Red;
            sheet.Cells["A1"].SetStyle(style);

            // Add a comment to cell A1
            int commentIndex = sheet.Comments.Add("A1");
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = "This is a test comment";

            // -----------------------------------------------------------------
            // Save with DisableDownlevelRevealedComments = true
            // This should disable downlevel-revealed conditional comments
            // while preserving other HTML features such as formulas, styles,
            // and comments (if ExportComments is enabled).
            // -----------------------------------------------------------------
            HtmlSaveOptions optionsDisable = new HtmlSaveOptions
            {
                DisableDownlevelRevealedComments = true,
                IsExportComments = true,          // ensure comments are exported
                ExportFormula = true,             // export formulas
                DisableCss = false,               // use external CSS to verify CSS handling
                HtmlVersion = HtmlVersion.Html5   // use HTML5 for modern output
            };
            workbook.Save("Output_DisableDownlevelRevealedComments_True.html", optionsDisable);

            // -----------------------------------------------------------------
            // Save with DisableDownlevelRevealedComments = false (default)
            // This serves as a baseline to compare that other features remain
            // identical except for the presence of downlevel-revealed comments.
            // -----------------------------------------------------------------
            HtmlSaveOptions optionsDefault = new HtmlSaveOptions
            {
                DisableDownlevelRevealedComments = false,
                IsExportComments = true,
                ExportFormula = true,
                DisableCss = false,
                HtmlVersion = HtmlVersion.Html5
            };
            workbook.Save("Output_DisableDownlevelRevealedComments_False.html", optionsDefault);

            // Inform the user that the files have been generated.
            Console.WriteLine("HTML files saved with DisableDownlevelRevealedComments set to true and false.");
            Console.WriteLine("Inspect the files to verify that other HTML features (styles, formulas, comments) are unchanged.");
        }
    }
}