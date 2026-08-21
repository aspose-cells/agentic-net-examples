// Title: Export Excel to Responsive HTML with Percentage Column Widths using Aspose.Cells for .NET
// Description: Load an Excel workbook with Aspose.Cells, enable HtmlSaveOptions.WidthScalable to output column widths as percentages, and save the file as responsive HTML.
// Keywords: Aspose.Cells | HtmlSaveOptions | WidthScalable | C# export Excel to HTML | percentage column widths | responsive HTML tables | convert .xlsx to HTML | Aspose.Cells .NET example
// Common Searches: Aspose.Cells WidthScalable C# example | export Excel to HTML with percentage column widths | how to make HTML output from Excel responsive using Aspose | save workbook as HTML Aspose.Cells WidthScalable true | convert .xlsx to responsive HTML .NET
// Developer Intent: Generate HTML from an Excel file where column widths are expressed as scalable percentages for responsive layouts.
// Use Cases: Create web‑ready reports that adapt to different screen sizes. | Automate batch conversion of multiple spreadsheets into responsive HTML tables for intranet portals. | Embed Excel‑derived tables in newsletters or documentation sites without fixed pixel widths.
// AI Prompts: Write C# code that converts an .xlsx file to HTML using Aspose.Cells with WidthScalable enabled and adds custom CSS for table styling. | Explain the effect of HtmlSaveOptions.WidthScalable on the generated HTML and how to adjust scaling behavior. | Provide a script that iterates through a directory of Excel files, converts each to responsive HTML with percentage column widths, and saves them to a target folder.

using System;
using Aspose.Cells;

// Load an Excel workbook with Aspose.Cells, enable HtmlSaveOptions.WidthScalable to output column widths as percentages, and save the file as responsive HTML.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to export column widths as scalable percentages
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.WidthScalable = true; // enables percentage column widths

        // Save the workbook as HTML using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}
