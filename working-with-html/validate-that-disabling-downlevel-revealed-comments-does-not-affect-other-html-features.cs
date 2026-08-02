// Title: Validate that disabling downlevel revealed comments in Aspose.Cells HTML export retains cell data and formatting (C#)
// Description: Creates a workbook with text, a number, bold blue styling, and a comment, then saves it to HTML twice—once with HtmlSaveOptions.DisableDownlevelRevealedComments enabled and once with the default setting. The HTML files are re‑loaded and the code verifies that cell values and the bold/blue style of A1 remain unchanged, outputting the validation results.
// Keywords: Aspose.Cells | HtmlSaveOptions | DisableDownlevelRevealedComments | HTML export | cell data integrity | cell formatting preservation | C# example | round‑trip validation | downlevel revealed comments | comment handling
// Common Searches: Aspose.Cells DisableDownlevelRevealedComments effect | HTML export cell values unchanged Aspose.Cells | preserve cell formatting when disabling downlevel comments | compare Aspose.Cells HTML output with and without downlevel comments | C# round‑trip test for Aspose.Cells HTML export
// Developer Intent: Confirm that enabling DisableDownlevelRevealedComments does not alter workbook data or styling after exporting to and importing from HTML.
// Use Cases: Export a spreadsheet to HTML while suppressing downlevel revealed comments and ensure numeric and string values stay intact after re‑import. | Verify that bold font and custom color applied to a cell survive the HTML round‑trip regardless of the comment setting. | Generate two HTML versions (with and without downlevel comments) and compare them to detect any unintended visual or data changes.
// AI Prompts: Write a C# unit test using Aspose.Cells that asserts cell values and formatting are identical after saving to HTML with DisableDownlevelRevealedComments enabled and then loading the file. | Generate C# code that loads two HTML files produced with HtmlSaveOptions (comments disabled vs. default) and programmatically reports differences in comments, styles, or data. | Explain how Aspose.Cells processes downlevel revealed comments during HTML export and why other workbook features such as cell values and formatting remain unaffected.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlDownlevelCommentValidation
{
    // Creates a workbook with text, a number, bold blue styling, and a comment, then saves it to HTML twice—once with HtmlSaveOptions.DisableDownlevelRevealedComments enabled and once with the default setting. The HTML files are re‑loaded and the code verifies that cell values and the bold/blue style of A1 remain unchanged, outputting the validation results.
    class Program
    {
        static void Main()
        {
            // 1. Create a sample workbook with data, formatting and a comment
            Workbook originalWorkbook = new Workbook();
            Worksheet sheet = originalWorkbook.Worksheets[0];

            // Add data
            sheet.Cells["A1"].PutValue("Hello World");
            sheet.Cells["B2"].PutValue(12345);

            // Apply formatting
            Style style = sheet.Cells["A1"].GetStyle();
            style.Font.IsBold = true;
            style.Font.Color = System.Drawing.Color.Blue;
            sheet.Cells["A1"].SetStyle(style);

            // Add a comment to A1
            int commentIndex = sheet.Comments.Add("A1");
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = "Sample comment";

            // 2. Save to HTML with DisableDownlevelRevealedComments = true
            HtmlSaveOptions optionsDisable = new HtmlSaveOptions();
            optionsDisable.DisableDownlevelRevealedComments = true;
            string htmlPathDisable = "output_disable.html";
            originalWorkbook.Save(htmlPathDisable, optionsDisable);

            // 3. Save to HTML with default setting (false) for comparison
            HtmlSaveOptions optionsDefault = new HtmlSaveOptions(); // default is false
            string htmlPathDefault = "output_default.html";
            originalWorkbook.Save(htmlPathDefault, optionsDefault);

            // 4. Load the HTML files back into workbooks
            Workbook loadedDisable = new Workbook(htmlPathDisable);
            Workbook loadedDefault = new Workbook(htmlPathDefault);

            // 5. Validate that cell values are unchanged
            bool valuesMatch = 
                loadedDisable.Worksheets[0].Cells["A1"].StringValue == "Hello World" &&
                loadedDisable.Worksheets[0].Cells["B2"].IntValue == 12345 &&
                loadedDefault.Worksheets[0].Cells["A1"].StringValue == "Hello World" &&
                loadedDefault.Worksheets[0].Cells["B2"].IntValue == 12345;

            // 6. Validate that formatting (bold & color) is preserved
            Style loadedStyle = loadedDisable.Worksheets[0].Cells["A1"].GetStyle();
            bool formattingMatch = loadedStyle.Font.IsBold && loadedStyle.Font.Color.ToArgb() == System.Drawing.Color.Blue.ToArgb();

            // 7. Output validation results
            Console.WriteLine("Cell values match after disabling downlevel comments: " + valuesMatch);
            Console.WriteLine("Cell formatting (bold & color) preserved: " + formattingMatch);
        }
    }
}
