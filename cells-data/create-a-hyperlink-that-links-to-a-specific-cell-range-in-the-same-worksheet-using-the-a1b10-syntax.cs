// Title: Aspose.Cells C# – Add an internal hyperlink to range A1:B10 on the same worksheet
// Description: Demonstrates how to create a new Workbook, add a hyperlink in cell C1 that points to the internal range A1:B10 using the "#" prefix, set custom display text, and save the file as HyperlinkToRange.xlsx with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | internal hyperlink | Excel range hyperlink | #A1:B10 | Worksheet.Hyperlinks.Add | navigate to range | hyperlink to cell range | Excel navigation link | Aspose.Cells example
// Common Searches: Aspose.Cells add hyperlink to same sheet range | C# internal hyperlink Excel Aspose | link to A1:B10 using Aspose.Cells | Worksheet.Hyperlinks.Add syntax | create navigation link in Excel with Aspose
// Developer Intent: Insert a hyperlink that jumps to a specific cell range within the current worksheet.
// Use Cases: Build a table of contents in an Excel report where each entry links to its data section. | Add quick‑access navigation cells in a dashboard to scroll to chart data ranges. | Create summary cells that open detailed tables when clicked, improving workbook usability.
// AI Prompts: Generate C# code with Aspose.Cells to add an internal hyperlink that points to range A1:B10 on the same worksheet. | Explain why the "#" prefix is required for internal references in Worksheet.Hyperlinks.Add. | Show how to customize the display text of a hyperlink and save the workbook using Aspose.Cells.

using System;
using Aspose.Cells;

namespace HyperlinkToRangeExample
{
    // Demonstrates how to create a new Workbook, add a hyperlink in cell C1 that points to the internal range A1:B10 using the "#" prefix, set custom display text, and save the file as HyperlinkToRange.xlsx with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a hyperlink in cell C1 that points to the range A1:B10 on the same sheet
            // The address uses the "#" prefix to indicate an internal reference
            worksheet.Hyperlinks.Add("C1", 1, 1, "#A1:B10");

            // Set the display text for the hyperlink
            worksheet.Cells["C1"].PutValue("Go to A1:B10");

            // Save the workbook
            workbook.Save("HyperlinkToRange.xlsx");
        }
    }
}
