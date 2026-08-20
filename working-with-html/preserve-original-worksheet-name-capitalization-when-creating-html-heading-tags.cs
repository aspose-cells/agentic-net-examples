// Title: Preserve Worksheet Name Capitalization in HTML Export with Aspose.Cells for .NET
// Description: This .NET example creates a workbook, assigns a mixed‑case sheet name, uses the &A placeholder in the page header, enables ExportPageHeaders, and saves the workbook as a single HTML file so the generated heading tag reflects the exact sheet name.
// Keywords: Aspose.Cells HTML export | worksheet name case preservation | ExportPageHeaders option | HTML header &A placeholder | C# preserve sheet name capitalization | single HTML file all sheets
// Common Searches: Aspose.Cells keep sheet name case in HTML | HTML export page header preserve original worksheet name | C# Aspose.Cells &A placeholder usage | ExportPageHeaders true example | Save multiple worksheets to one HTML file Aspose
// Developer Intent: I need the HTML output to show the worksheet title with the same capitalization used in the workbook.
// Use Cases: Creating web‑ready reports where sheet titles must match corporate naming standards | Building printable HTML dashboards that retain case‑sensitive identifiers | Consolidating several worksheets into one HTML document while preserving each sheet’s exact name
// AI Prompts: Generate code that exports each worksheet to its own HTML file while keeping the original case of the sheet names. | Explain how the &A placeholder works and why it does not alter capitalization during HTML export. | Show how to customize the HTML header format (font, alignment) without losing the sheet name case. | Provide a step‑by‑step guide to enable ExportPageHeaders and SaveAsSingleFile for multi‑sheet workbooks.

using System;
using Aspose.Cells;

namespace PreserveWorksheetNameCapitalization
{
    // This .NET example creates a workbook, assigns a mixed‑case sheet name, uses the &A placeholder in the page header, enables ExportPageHeaders, and saves the workbook as a single HTML file so the generated heading tag reflects the exact sheet name.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set a custom worksheet name with mixed capitalization
            sheet.Name = "SalesData2021";

            // Add some sample data (optional, just to have content)
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(150);

            // Configure the page header to display the worksheet name.
            // Using the &A placeholder preserves the original capitalization of the sheet name.
            sheet.PageSetup.SetHeader(0, "&A"); // Left section of the header

            // Set HTML save options to export page headers.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportPageHeaders = true,   // Enable exporting of page headers
                SaveAsSingleFile = true,    // Save all sheets into a single HTML file
                ShowAllSheets = true        // Ensure all worksheets are rendered
            };

            // Save the workbook as HTML. The generated <h1> (or similar) tag will contain
            // the worksheet name exactly as defined ("SalesData2021").
            workbook.Save("PreservedWorksheetName.html", htmlOptions);
        }
    }
}
