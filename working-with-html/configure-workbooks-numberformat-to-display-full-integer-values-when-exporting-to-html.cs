// Title: Set a workbook’s default style to an integer‑only number format and export it as HTML with Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells in C# to assign the custom number format "0" to the workbook’s DefaultStyle and then save the workbook as HTML. | Configure HtmlSaveOptions after applying a whole‑number display format to all cells to generate an HTML file without decimal or scientific notation.
// Common Searches: Aspose.Cells C# export Excel to HTML with integer only number format | How to set default number format to 0 for HTML conversion using Aspose.Cells .NET | Prevent decimals and scientific notation when saving workbook as HTML with Aspose.Cells | Apply custom number format to entire workbook before HTML export in Aspose.Cells | C# Aspose.Cells default style number format for HTML output
// Tags: default style custom number format Aspose.Cells | HTML export integer number format C# | Aspose.Cells prevent scientific notation HTML | set workbook default style Aspose.Cells .NET | custom number format 0 Aspose.Cells

using Aspose.Cells;
using System;

// The example loads an Excel workbook, changes its DefaultStyle to use the custom number format "0" (displaying whole integers only), and then saves the workbook as an HTML file with HtmlSaveOptions, ensuring numbers appear without decimals or scientific notation.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Configure the default style to use a custom number format that shows full integers
        Style defaultStyle = workbook.DefaultStyle;
        defaultStyle.Custom = "0"; // Displays numbers without decimals or scientific notation
        workbook.DefaultStyle = defaultStyle;

        // Save the workbook as HTML with the applied number format
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        workbook.Save("output.html", htmlOptions);
    }
}
