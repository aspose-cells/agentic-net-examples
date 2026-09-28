// Title: Export an Aspose.Cells workbook to HTML with each worksheet name displayed as a heading (C#)
// AI Prompts: Write C# code that creates a workbook with multiple worksheets, populates cells, and saves it to a single HTML file where each worksheet name appears as an <h1> heading by setting HtmlSaveOptions.ExportHeadings to true. | Update an existing Aspose.Cells C# example to enable ExportHeadings so the generated HTML includes the sheet titles as headings. | Provide a concise C# snippet that demonstrates adding two sheets, inserting data, and exporting to HTML with sheet names rendered as headings.
// Common Searches: Aspose.Cells HtmlSaveOptions ExportHeadings property C# example | how to include worksheet names as headings when saving Excel to HTML with Aspose.Cells | C# convert workbook with multiple sheets to HTML with sheet titles
// Tags: Aspose.Cells export HTML with worksheet headings | HtmlSaveOptions ExportHeadings C# | multiple worksheets to HTML Aspose.Cells | C# generate HTML from Excel with sheet titles | save workbook as HTML using Aspose.Cells

using Aspose.Cells;
using System;

// Shows how to create a workbook with two worksheets, fill cells, enable HtmlSaveOptions.ExportHeadings, and save the workbook as an HTML file where each worksheet name is rendered as a heading.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // First worksheet: set name and some data
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "FirstSheet";
        sheet1.Cells["A1"].PutValue("Hello");
        sheet1.Cells["B2"].PutValue(123);

        // Add a second worksheet with its own data
        Worksheet sheet2 = workbook.Worksheets.Add("SecondSheet");
        sheet2.Cells["A1"].PutValue("World");
        sheet2.Cells["B2"].PutValue(456);

        // Configure HTML save options to include worksheet names as headings
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportHeadings = true; // Export worksheet names as headings

        // Save the workbook as an HTML file using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}
