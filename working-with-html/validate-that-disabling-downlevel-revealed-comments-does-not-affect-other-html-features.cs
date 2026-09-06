// Title: Verify that disabling downlevel revealed comments during HTML export does not affect cell values, formulas, or formatting with Aspose.Cells for .NET
// AI Prompts: Create C# code that builds a workbook, saves it to HTML with HtmlSaveOptions.DisableDownlevelRevealedComments set to true, reloads the HTML file, and asserts that the original cell values, formulas, and styles are unchanged. | Write a C# unit test using Aspose.Cells that confirms disabling downlevel revealed comments in HTML export preserves data integrity, formula accuracy, and header styling. | Show how to compare the style and formula of cells in the original workbook with those in a workbook loaded from the generated HTML after turning off downlevel revealed comments.
// Common Searches: Aspose.Cells C# verify HTML export keeps formatting when downlevel revealed comments are disabled | How to test that formulas remain after saving workbook to HTML with DisableDownlevelRevealedComments | C# load HTML saved by Aspose.Cells and compare cell values and styles | Impact of disabling downlevel revealed comments on Aspose.Cells HTML output
// Tags: Aspose.Cells HTML export disable downlevel comments | preserve cell formatting during HTML conversion | validate formula retention after HTML save | C# workbook reload from generated HTML | check cell values after disabling downlevel revealed comments

using System;
using System.Drawing;
using Aspose.Cells;

// The program creates a workbook, adds numeric data, a SUM formula, and bold blue header styling, saves it to HTML with downlevel revealed comments disabled, reloads the HTML into a new workbook, and verifies that cell values, the formula, and the header formatting remain unchanged.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Data";

        // Populate cells with data and a formula
        sheet.Cells["A1"].PutValue("Number");
        sheet.Cells["A2"].PutValue(10);
        sheet.Cells["A3"].PutValue(20);
        sheet.Cells["B1"].PutValue("Sum");
        sheet.Cells["B2"].Formula = "SUM(A2:A3)";

        // Apply simple formatting (bold blue font) to header cells
        Style headerStyle = workbook.CreateStyle();
        headerStyle.Font.Color = Color.Blue;
        headerStyle.Font.IsBold = true;
        sheet.Cells["A1"].SetStyle(headerStyle);
        sheet.Cells["B1"].SetStyle(headerStyle);

        // Save the workbook to HTML with downlevel revealed comments disabled
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        htmlOptions.DisableDownlevelRevealedComments = true; // Disable the feature
        string htmlPath = "output.html";
        workbook.Save(htmlPath, htmlOptions);

        // Load the generated HTML back into a new workbook
        Workbook loadedWorkbook = new Workbook(htmlPath);
        Worksheet loadedSheet = loadedWorkbook.Worksheets[0];

        // Validate that cell values are unchanged
        bool valuesMatch = loadedSheet.Cells["A2"].StringValue == "10" &&
                           loadedSheet.Cells["A3"].StringValue == "20";

        // Validate that the formula is preserved
        bool formulaPreserved = loadedSheet.Cells["B2"].Formula == "SUM(A2:A3)";

        // Validate that formatting (bold and blue font) is still applied
        Style loadedStyle = loadedSheet.Cells["A1"].GetStyle();
        bool formattingMatch = loadedStyle.Font.IsBold &&
                               loadedStyle.Font.Color.ToArgb() == Color.Blue.ToArgb();

        // Output validation results
        Console.WriteLine($"Values match: {valuesMatch}");
        Console.WriteLine($"Formula preserved: {formulaPreserved}");
        Console.WriteLine($"Formatting preserved: {formattingMatch}");
    }
}
