// Title: Load HTML with DIV layout, edit a cell, and save as inline‑style HTML using Aspose.Cells for .NET
// Description: Demonstrates how to load an HTML file while preserving its <div> structure (SupportDivTag), change a worksheet cell, and export the workbook back to HTML with only inline styles (DisableCss).
// Keywords: Aspose.Cells | C# | .NET | load HTML with div layout | SupportDivTag | modify cell | save HTML inline styles | DisableCss | HTML to spreadsheet conversion | export spreadsheet to HTML
// Common Searches: Aspose.Cells preserve DIV layout when loading HTML | How to disable CSS generation in Aspose.Cells HTML export | Change cell value after loading HTML with Aspose.Cells | C# load HTML workbook and save with inline styles | SupportDivTag option example
// Developer Intent: Load an HTML document, keep its DIV‑based layout, update a cell, and re‑export to HTML using only inline CSS.
// Use Cases: Convert a DIV‑based web template into an editable spreadsheet, modify data programmatically, and generate email‑ready HTML with inline styling. | Automate report generation where a pre‑designed HTML layout is loaded, dynamic values are inserted into specific cells, and the final output is saved without external CSS files.
// AI Prompts: Generate C# code with Aspose.Cells to load an HTML file using SupportDivTag, change cell B2, and save the workbook as HTML with DisableCss enabled. | Explain the effects of SupportDivTag and DisableCss on HTML import/export in Aspose.Cells for .NET, including performance and styling considerations.

using System;
using Aspose.Cells;

// Demonstrates how to load an HTML file while preserving its <div> structure (SupportDivTag), change a worksheet cell, and export the workbook back to HTML with only inline styles (DisableCss).
class Program
{
    static void Main()
    {
        // Path to the source HTML file that contains <div> layout
        string inputHtmlPath = "input.html";

        // Load the HTML file with SupportDivTag enabled to preserve <div> layout
        HtmlLoadOptions loadOptions = new HtmlLoadOptions(LoadFormat.Html);
        loadOptions.SupportDivTag = true; // preserve DIV based layout
        Workbook workbook = new Workbook(inputHtmlPath, loadOptions);

        // Modify a cell in the first worksheet (e.g., set A1 value)
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Modified Value");

        // Prepare HTML save options with DisableCss = true to use only inline styles
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.DisableCss = true; // no external CSS, inline styles only

        // Save the workbook back to HTML
        string outputHtmlPath = "output.html";
        workbook.Save(outputHtmlPath, saveOptions);

        Console.WriteLine("HTML loaded, cell modified, and saved with inline styles.");
    }
}
