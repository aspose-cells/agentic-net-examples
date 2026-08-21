// Title: Print Gridlines and Hide Row/Column Headers in Aspose.Cells (C#)
// Description: Shows how to enable gridline printing and suppress row and column header display for a worksheet using Aspose.Cells for .NET, then save the workbook.
// Keywords: Aspose.Cells | C# | print gridlines | hide row headers | hide column headers | worksheet page setup | gridlines without headings | Excel printing options | IsRowColumnHeadersVisible | PageSetup.PrintGridlines
// Common Searches: Aspose.Cells print gridlines C# | hide row and column headings Aspose.Cells | gridlines without headers Excel .NET | PageSetup.PrintGridlines example | IsRowColumnHeadersVisible usage
// Developer Intent: Configure a worksheet so printed pages show gridlines but omit row and column header labels.
// Use Cases: Create printable reports with a clean grid‑only layout. | Generate Excel templates where headers are hidden for end‑user printing. | Automate workbook creation for dashboards that require only data cells and gridlines.
// AI Prompts: Write C# code using Aspose.Cells to enable gridline printing and hide row/column headers for the active worksheet, then save the file. | Show how to set PageSetup.PrintGridlines and IsRowColumnHeadersVisible properties in Aspose.Cells to produce a workbook that prints only gridlines. | Provide an example that adds sample data, configures printing options to show gridlines without headings, and saves the workbook.

using System;
using Aspose.Cells;

// Shows how to enable gridline printing and suppress row and column header display for a worksheet using Aspose.Cells for .NET, then save the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet (current worksheet)
        Worksheet worksheet = workbook.Worksheets[0];

        // Enable printing of gridlines on the page
        worksheet.PageSetup.PrintGridlines = true;

        // Disable display of row and column headings
        worksheet.IsRowColumnHeadersVisible = false;

        // (Optional) Add some sample data to visualize the settings
        worksheet.Cells["A1"].PutValue("Sample Data");
        worksheet.Cells["B2"].PutValue(123);

        // Save the workbook
        workbook.Save("Gridlines_NoHeadings.xlsx");
    }
}
