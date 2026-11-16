// Title: Export a worksheet to HTML with a frozen top row using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export worksheet to HTML with frozen header row | how to keep frozen panes visible when saving Excel as HTML using Aspose.Cells | C# FreezePanes then HtmlSaveOptions ExportActiveWorksheetOnly example | save Excel file as HTML with frozen top row in Aspose.Cells for .NET | preserve frozen rows in HTML output using Aspose.Cells C#
// Tags: Aspose.Cells FreezePanes HTML export | HtmlSaveOptions ExportActiveWorksheetOnly C# | frozen rows retained in HTML conversion | export active worksheet to HTML Aspose.Cells | C# freeze top row before HTML conversion

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, adds sample data, freezes the first row with FreezePanes(1,0,0,0), configures HtmlSaveOptions.ExportActiveWorksheetOnly = true, and saves the workbook as an HTML file where the frozen header row stays visible while scrolling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (optional, for demonstration)
            sheet.Cells["A1"].PutValue("Header 1");
            sheet.Cells["B1"].PutValue("Header 2");
            sheet.Cells["A2"].PutValue("Row 1, Col 1");
            sheet.Cells["B2"].PutValue("Row 1, Col 2");
            sheet.Cells["A3"].PutValue("Row 2, Col 1");
            sheet.Cells["B3"].PutValue("Row 2, Col 2");

            // Freeze the top row so that headers stay visible while scrolling
            // Overload with four parameters is used for compatibility with newer API versions
            sheet.FreezePanes(1, 0, 0, 0);

            // Set HTML save options to export only the active worksheet
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportActiveWorksheetOnly = true
                // Freeze panes are exported by default; no explicit property needed
            };

            // Define output file path
            string outputPath = "FrozenHeaders.html";

            // Save the workbook as HTML
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"HTML file saved successfully to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
