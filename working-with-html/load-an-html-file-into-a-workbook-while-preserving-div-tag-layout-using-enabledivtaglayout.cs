// Title: C# – Load HTML into Aspose.Cells Workbook while Preserving DIV Layout (HtmlLoadOptions.SupportDivTag)
// Description: Demonstrates how to create HtmlLoadOptions with SupportDivTag enabled, load an HTML file (e.g., input.html) into an Aspose.Cells Workbook, verify cell content, and save the result as an XLSX file, ensuring the original DIV‑based page layout is retained.
// Keywords: Aspose.Cells | HtmlLoadOptions | SupportDivTag | C# HTML to Excel | preserve DIV layout | convert web page to Excel | load HTML workbook | Excel export .NET | DIV‑based layout conversion | Aspose.Cells example
// Common Searches: Aspose.Cells preserve div layout when loading html | HtmlLoadOptions SupportDivTag C# example | Convert HTML with divs to Excel using Aspose.Cells | Load HTML file into workbook keeping CSS divs | How to keep div positioning in Excel export Aspose
// Developer Intent: Load an HTML document into an Aspose.Cells workbook and retain the layout defined by DIV tags.
// Use Cases: Transform a web page that relies on DIV containers into an Excel file without losing element positioning. | Generate Excel reports from HTML templates that use DIV‑based layouts for precise formatting. | Validate that specific cell values are correctly imported from an HTML source containing DIV structures.
// AI Prompts: Provide C# code that loads an HTML file with DIV layout into an Aspose.Cells workbook using HtmlLoadOptions.SupportDivTag and saves it as XLSX. | Explain how to configure HtmlLoadOptions to keep DIV tag positioning when converting HTML to Excel with Aspose.Cells. | Show an example of reading cell values after importing an HTML file that uses DIV tags into a workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsDivTagDemo
{
    // Demonstrates how to create HtmlLoadOptions with SupportDivTag enabled, load an HTML file (e.g., input.html) into an Aspose.Cells Workbook, verify cell content, and save the result as an XLSX file, ensuring the original DIV‑based page layout is retained.
    class Program
    {
        static void Main()
        {
            // Create HTML load options and enable DIV tag layout support
            HtmlLoadOptions loadOptions = new HtmlLoadOptions(LoadFormat.Html);
            loadOptions.SupportDivTag = true;

            // Load the HTML file into a workbook using the specified options
            Workbook workbook = new Workbook("input.html", loadOptions);

            // (Optional) Access a cell to verify that content was loaded
            Worksheet sheet = workbook.Worksheets[0];
            Console.WriteLine("Cell A1 value: " + sheet.Cells["A1"].StringValue);

            // Save the workbook to an Excel file
            workbook.Save("output.xlsx");
        }
    }
}
