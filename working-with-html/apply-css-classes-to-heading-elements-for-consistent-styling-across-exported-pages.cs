// Title: Export an Excel workbook to HTML with a bold gray header row rendered as <th> elements using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a workbook, formats the first row with bold font and a light‑gray background, enables ExportRowColumnHeadings in HtmlSaveOptions, and saves the workbook as an HTML file. | Show how to assign a custom CSS class to the generated <th> elements when exporting an Excel sheet to HTML with Aspose.Cells in .NET.
// Common Searches: Aspose.Cells C# export Excel to HTML with header row as <th> tags | how to use ExportRowColumnHeadings in HtmlSaveOptions for .NET | apply bold and gray style to Excel header row before HTML conversion using Aspose.Cells | set custom CSS class for table header cells in Aspose.Cells HTML output | C# example of saving a workbook as HTML with styled heading cells
// Tags: Aspose.Cells HTML export header row styling | C# ExportRowColumnHeadings property usage | custom CSS class for HTML table headers Aspose.Cells | Excel heading row styling for HTML export | generate HTML with <th> elements from Excel workbook

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a new workbook, writes header values to the first row, applies a bold font and light‑gray background style, enables ExportRowColumnHeadings so the headers are emitted as <th> elements, optionally assigns a CSS class to those headers, and saves the workbook as an HTML file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate heading cells (first row)
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["C1"].PutValue("Quantity");

            // Define a style for the heading row
            Style headingStyle = workbook.CreateStyle();
            headingStyle.Font.IsBold = true;
            headingStyle.ForegroundColor = Color.LightGray;
            headingStyle.Pattern = BackgroundType.Solid;

            // Apply the style to the heading range
            Aspose.Cells.Range headingRange = sheet.Cells.CreateRange("A1:C1");
            headingRange.ApplyStyle(headingStyle, new StyleFlag() { All = true });

            // Configure HTML export options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Export row/column headings as <th> elements (new property name)
            htmlOptions.ExportRowColumnHeadings = true;

            // Define output file path
            string outputPath = "ExportedWithHeadings.html";

            // Save the workbook as an HTML file with the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully exported to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
