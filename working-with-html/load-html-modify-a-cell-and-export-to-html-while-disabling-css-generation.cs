// Title: Load HTML, modify a cell, and export with inline styles using Aspose.Cells for .NET
// Description: Shows how to open an HTML workbook with Aspose.Cells, change a cell's value, turn off external CSS generation, and save the file as self‑contained HTML that uses only inline styling.
// Keywords: Aspose.Cells load HTML | modify cell C# | HtmlSaveOptions DisableCss | export HTML inline styles | C# spreadsheet to HTML
// Common Searches: Aspose.Cells load HTML file and edit cells | disable CSS when saving workbook to HTML Aspose.Cells | C# save HTML with inline styles only Aspose.Cells | how to edit HTML spreadsheet with Aspose.Cells | inline CSS export Aspose.Cells .NET
// Developer Intent: Open an HTML workbook, update a cell, and save it as HTML that contains only inline CSS.
// Use Cases: Refresh data in an HTML‑based spreadsheet without creating separate style sheets. | Create email‑compatible HTML reports where all formatting must be inline. | Migrate legacy HTML spreadsheets to updated content while keeping the visual layout self‑contained.
// AI Prompts: Generate C# code that loads an HTML file into Aspose.Cells, changes cell B2, and saves the workbook with DisableCss enabled. | Explain the effect of HtmlSaveOptions.DisableCss in Aspose.Cells and provide a short example. | Suggest a method to batch‑process multiple HTML workbooks, modify a specific cell in each, and export them with only inline styles using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Shows how to open an HTML workbook with Aspose.Cells, change a cell's value, turn off external CSS generation, and save the file as self‑contained HTML that uses only inline styling.
class Program
{
    static void Main()
    {
        // Load the workbook from an existing HTML file
        Workbook workbook = new Workbook("input.html");

        // Modify a cell (example: set value of cell A1)
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Modified Value");

        // Configure HTML save options to use only inline styles (disable CSS files)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.DisableCss = true;

        // Save the workbook back to HTML with the specified options
        workbook.Save("output.html", htmlOptions);
    }
}
