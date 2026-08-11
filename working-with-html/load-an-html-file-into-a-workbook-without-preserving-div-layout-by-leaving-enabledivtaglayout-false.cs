// Title: Load HTML into an Aspose.Cells Workbook without preserving DIV layout (SupportDivTag = false) – C#
// Description: Demonstrates how to use Aspose.Cells HtmlLoadOptions with SupportDivTag set to false to import an HTML file into a Workbook, ignore <div> tag positioning, read a cell for verification, and save the result as an XLSX file.
// Keywords: Aspose.Cells HTML import | SupportDivTag false | disable DIV layout Aspose.Cells | C# load HTML to Excel | HtmlLoadOptions example | convert HTML to XLSX | ignore div tags Excel conversion
// Common Searches: Aspose.Cells load HTML without div layout | HtmlLoadOptions SupportDivTag property C# example | convert HTML to Excel ignoring div tags | how to disable div tag layout when loading HTML in Aspose.Cells | C# Aspose.Cells HTML to workbook without preserving divs
// Developer Intent: Import an HTML document into an Aspose.Cells Workbook while disabling the preservation of <div> tag layout.
// Use Cases: Transform simple web reports into Excel sheets without extra spacing caused by DIV elements. | Extract tabular data from HTML emails where DIV formatting is irrelevant. | Batch‑process large numbers of HTML files to Excel, improving speed by skipping DIV layout handling.
// AI Prompts: Provide C# code that loads an HTML file into an Aspose.Cells Workbook with SupportDivTag set to false and saves it as XLSX. | Show how to load multiple HTML files, disable DIV layout, and place each file on a separate worksheet in one workbook using Aspose.Cells. | Explain the effect of the SupportDivTag property on HTML‑to‑Excel conversion and how to confirm that DIV layout is not retained.

using System;
using Aspose.Cells;

// Demonstrates how to use Aspose.Cells HtmlLoadOptions with SupportDivTag set to false to import an HTML file into a Workbook, ignore <div> tag positioning, read a cell for verification, and save the result as an XLSX file.
class LoadHtmlWithoutDivLayout
{
    static void Main()
    {
        // Create HtmlLoadOptions; the default SupportDivTag is false,
        // which means the layout of <div> tags will not be preserved.
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();
        loadOptions.SupportDivTag = false; // explicit for clarity

        // Load the HTML file into a workbook using the specified options.
        Workbook workbook = new Workbook("input.html", loadOptions);

        // Example: read a cell value to verify the load succeeded.
        Worksheet sheet = workbook.Worksheets[0];
        Console.WriteLine("A1 value: " + sheet.Cells["A1"].StringValue);

        // Save the workbook to an Excel file.
        workbook.Save("output.xlsx");
    }
}
