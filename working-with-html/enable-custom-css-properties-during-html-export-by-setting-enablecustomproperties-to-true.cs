// Title: Enable CSS Custom Properties in HTML Export with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, set HtmlSaveOptions.EnableCssCustomProperties to true, and save the file as HTML so the output uses CSS custom properties (variables) for styling.
// Keywords: Aspose.Cells HTML export | EnableCssCustomProperties | CSS custom properties Aspose | HtmlSaveOptions C# | Excel to HTML with CSS variables | Aspose.Cells .NET example
// Common Searches: Aspose.Cells enable CSS custom properties | HtmlSaveOptions EnableCssCustomProperties example | Export Excel to HTML with CSS variables C# | How to use CSS custom properties in Aspose.Cells HTML output | Aspose.Cells HTML export lightweight styling
// Developer Intent: Activate CSS custom properties when converting an Excel workbook to HTML using Aspose.Cells.
// Use Cases: Produce compact HTML reports where repeated styles are stored as reusable CSS variables. | Create web‑ready spreadsheet snapshots that support dynamic theming via CSS custom properties. | Reduce HTML file size for large workbooks by consolidating style definitions into variables.
// AI Prompts: Show a C# snippet that saves an Aspose.Cells workbook to HTML with EnableCssCustomProperties set to true and illustrates how the generated CSS variables can be referenced. | Explain the performance and maintenance benefits of using EnableCssCustomProperties in Aspose.Cells HTML export. | Provide code to export a workbook as HTML with custom CSS properties enabled and link an external stylesheet that utilizes those variables.

using System;
using Aspose.Cells;

namespace AsposeCellsCustomCssDemo
{
    // Demonstrates how to create a workbook, set HtmlSaveOptions.EnableCssCustomProperties to true, and save the file as HTML so the output uses CSS custom properties (variables) for styling.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");

            // Create HTML save options and enable CSS custom properties
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.EnableCssCustomProperties = true; // Optimize HTML using CSS custom properties

            // Save the workbook as HTML with the specified options
            workbook.Save("OutputWithCustomCss.html", htmlOptions);

            Console.WriteLine("HTML file saved with EnableCssCustomProperties set to true.");
        }
    }
}
