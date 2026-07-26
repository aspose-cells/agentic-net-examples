// Title: C# – Load HTML into Aspose.Cells Workbook without DIV layout (SupportDivTag = false)
// Description: Shows how to use Aspose.Cells for .NET to import an HTML file into a Workbook while disabling DIV tag layout. The sample creates HtmlLoadOptions, sets SupportDivTag to false, loads "input.html", reads cell A1 from the first worksheet, and saves the workbook as "output.xlsx".
// Keywords: Aspose.Cells | HtmlLoadOptions | SupportDivTag | C# | .NET | HTML to Excel | disable div layout | convert HTML to XLSX | import HTML without DIV tags
// Common Searches: Aspose.Cells load HTML without div layout C# | HtmlLoadOptions SupportDivTag false example | Convert HTML to Excel ignoring DIV tags | C# import HTML to XLSX Aspose.Cells
// Developer Intent: Import an HTML document into a Workbook while ignoring the layout of <div> elements.
// Use Cases: Transform web‑based reports that rely on DIV spacing into clean Excel sheets. | Automate batch conversion of HTML newsletters to XLSX while stripping DIV structures. | Generate Excel data extracts from HTML snippets where layout tags are irrelevant.
// AI Prompts: Generate C# code using Aspose.Cells to load an HTML file with SupportDivTag set to false and save it as XLSX. | Explain how the SupportDivTag property influences worksheet structure when importing HTML with Aspose.Cells. | Provide a step‑by‑step guide to batch convert a folder of HTML files to Excel while disabling DIV layout using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlImportExample
{
    // Shows how to use Aspose.Cells for .NET to import an HTML file into a Workbook while disabling DIV tag layout. The sample creates HtmlLoadOptions, sets SupportDivTag to false, loads "input.html", reads cell A1 from the first worksheet, and saves the workbook as "output.xlsx".
    class Program
    {
        static void Main()
        {
            // Create HTML load options. The default value of SupportDivTag is false,
            // which means the layout of <div> tags will not be preserved.
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            loadOptions.SupportDivTag = false; // optional, shown for clarity

            // Load the HTML file into a workbook using the specified options.
            Workbook workbook = new Workbook("input.html", loadOptions);

            // Access the first worksheet (optional, just to demonstrate usage).
            Worksheet sheet = workbook.Worksheets[0];
            Console.WriteLine("First cell value: " + sheet.Cells["A1"].StringValue);

            // Save the workbook to an Excel file.
            workbook.Save("output.xlsx");
        }
    }
}
