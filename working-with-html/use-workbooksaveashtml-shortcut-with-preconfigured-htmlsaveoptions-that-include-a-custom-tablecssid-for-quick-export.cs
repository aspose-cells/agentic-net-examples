// Title: Export Workbook to HTML with a Custom TableCssId Using Aspose.Cells HtmlSaveOptions (C#)
// Description: Shows how to create an Aspose.Cells workbook, assign a custom TableCssId via HtmlSaveOptions, and save the file as HTML in .NET. The example uses a pre‑configured HtmlSaveOptions instance for a fast, reusable export with targeted CSS styling.
// Keywords: Aspose.Cells | HtmlSaveOptions | TableCssId | C# | export workbook to HTML | custom CSS ID | Workbook.SaveAsHtml | .NET | Excel to HTML conversion | pre‑configured options | quick HTML export
// Common Searches: Aspose.Cells set TableCssId when exporting to HTML | C# export Excel workbook as HTML with custom CSS identifier | HtmlSaveOptions TableCssId example | How to use pre‑configured HtmlSaveOptions for HTML export | Aspose.Cells quick HTML export .NET
// Developer Intent: Export an Excel workbook to HTML while applying a custom CSS ID to the generated table for easy styling.
// Use Cases: Apply a specific stylesheet to the exported HTML table by defining HtmlSaveOptions.TableCssId. | Reuse a single HtmlSaveOptions object across multiple workbook exports to maintain consistent table styling. | Integrate HTML export into a web app that references a CSS rule targeting the custom table ID.
// AI Prompts: Generate C# code that creates a workbook, fills cells, and saves it as HTML with a custom TableCssId using Aspose.Cells. | Show how to load an existing workbook and export it to HTML with a pre‑configured HtmlSaveOptions that includes TableCssId and additional formatting options. | Provide an example of linking an external CSS file to the HTML output by using the TableCssId defined in HtmlSaveOptions.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Shows how to create an Aspose.Cells workbook, assign a custom TableCssId via HtmlSaveOptions, and save the file as HTML in .NET. The example uses a pre‑configured HtmlSaveOptions instance for a fast, reusable export with targeted CSS styling.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data
            worksheet.Cells["A1"].PutValue("Name");
            worksheet.Cells["B1"].PutValue("Age");
            worksheet.Cells["A2"].PutValue("John");
            worksheet.Cells["B2"].PutValue(30);
            worksheet.Cells["A3"].PutValue("Alice");
            worksheet.Cells["B3"].PutValue(25);

            // Configure HTML save options with a custom TableCssId
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            htmlOptions.TableCssId = "custom-table-style";

            // Save the workbook as HTML using the pre‑configured options
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("Workbook exported to HTML with TableCssId = " + htmlOptions.TableCssId);
        }
    }
}
