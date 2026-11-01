// Title: Force full integer display when exporting a workbook to HTML with Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a style with Custom = "0" to show integers without scientific notation, applies it to a cell range, and saves the workbook as HTML using Aspose.Cells. | Write a method that sets the NumberFormat of selected cells to an integer‑only format and then exports the workbook to HTML in Aspose.Cells .NET.
// Common Searches: Aspose.Cells C# export to HTML keep integer formatting | prevent scientific notation in HTML output using Aspose.Cells | set custom number format 0 for cells before saving as HTML Aspose.Cells | apply integer number format to a range in Aspose.Cells .NET | how to force full integer values in Aspose.Cells HTML export
// Tags: integer number format Aspose.Cells | HTML export number format .NET | apply style to cell range Aspose.Cells | custom number format 0 Aspose.Cells | prevent scientific notation HTML Aspose.Cells

using System;
using Aspose.Cells;

// Creates a workbook, writes integer values, defines a style with Custom = "0" to force integer‑only display, applies the style to cells A1:A3, and saves the workbook as an HTML file.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some integer values
        sheet.Cells["A1"].PutValue(123456);
        sheet.Cells["A2"].PutValue(987654321);
        sheet.Cells["A3"].PutValue(-42);

        // Define a style that forces full integer display (no scientific notation)
        Style intStyle = workbook.CreateStyle();
        // Custom number format "0" displays the number as an integer without decimals
        intStyle.Custom = "0";

        // Apply the style to the range containing the integers
        StyleFlag flag = new StyleFlag();
        flag.NumberFormat = true; // Apply only the number format part of the style
        sheet.Cells.CreateRange("A1:A3").ApplyStyle(intStyle, flag);

        // Export the workbook to HTML
        workbook.Save("output.html", SaveFormat.Html);
    }
}
