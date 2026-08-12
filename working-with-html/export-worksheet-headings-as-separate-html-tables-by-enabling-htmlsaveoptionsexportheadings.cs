// Title: Export Row and Column Headings as Separate HTML Tables with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to enable HtmlSaveOptions.ExportHeadings in Aspose.Cells to save an Excel worksheet as HTML where row and column headings are rendered in their own tables.
// Keywords: Aspose.Cells ExportHeadings | C# HTML export Excel headings | separate heading tables Aspose | HtmlSaveOptions ExportHeadings example | Aspose.Cells .NET HTML output | export Excel column headings to HTML | row headings separate HTML table
// Common Searches: Aspose.Cells ExportHeadings C# example | how to save Excel headings as separate HTML tables | HtmlSaveOptions ExportHeadings property usage | export Excel worksheet to HTML with headings split | C# Aspose.Cells HTML export row and column headings
// Developer Intent: Generate an HTML file from a workbook where the worksheet’s row and column headings are placed in distinct tables by setting ExportHeadings to true.
// Use Cases: Design web reports that need independent styling for header rows and columns. | Improve accessibility by isolating headings for screen‑reader navigation. | Create printable HTML layouts where headings are positioned separately from data.
// AI Prompts: Show how to load an existing workbook, enable ExportHeadings, and save it as HTML with a custom file name. | Provide a snippet that adds CSS classes to the heading tables produced by ExportHeadings. | Explain the interaction between ExportHeadings and other HtmlSaveOptions such as ExportImagesAsBase64.

using System;
using Aspose.Cells;

namespace ExportHeadingsExample
{
    // Demonstrates how to enable HtmlSaveOptions.ExportHeadings in Aspose.Cells to save an Excel worksheet as HTML where row and column headings are rendered in their own tables.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["B2"].PutValue("Item");
            sheet.Cells["C2"].PutValue("Quantity");
            sheet.Cells["B3"].PutValue("Apples");
            sheet.Cells["C3"].PutValue(10);
            sheet.Cells["B4"].PutValue("Oranges");
            sheet.Cells["C4"].PutValue(20);

            // Configure HTML save options to export row/column headings
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.ExportHeadings = true; // enables separate tables for headings

            // Save the workbook as HTML
            workbook.Save("ExportHeadings.html", saveOptions);

            Console.WriteLine("Workbook saved with headings exported as separate HTML tables.");
        }
    }
}
