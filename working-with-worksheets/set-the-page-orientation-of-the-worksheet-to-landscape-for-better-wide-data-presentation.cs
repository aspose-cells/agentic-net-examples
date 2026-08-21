// Title: C# Aspose.Cells – Set Worksheet Page Orientation to Landscape
// Description: Shows how to create a workbook, select the first worksheet, and set its PageSetup to landscape mode for wide‑format printing or PDF export with Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# landscape orientation | set worksheet page orientation .NET | Excel print layout landscape Aspose | PageSetup orientation Aspose.Cells | landscape printing Excel C#
// Common Searches: Aspose.Cells set worksheet to landscape C# | change Excel page orientation programmatically .NET | print Excel sheet in landscape using Aspose | C# Aspose.Cells page setup orientation | landscape PDF export Aspose.Cells
// Developer Intent: Configure a worksheet to print in landscape layout.
// Use Cases: Print wide tables without column truncation. | Generate landscape‑formatted PDF reports directly from Excel. | Create printable charts that require a horizontal page layout. | Prepare invoices or schedules that need a horizontal orientation.
// AI Prompts: Provide C# code to set landscape orientation and adjust margins for a worksheet with Aspose.Cells. | Show how to toggle between portrait and landscape automatically based on column count. | Explain how to apply landscape orientation to all sheets in an existing workbook using Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to create a workbook, select the first worksheet, and set its PageSetup to landscape mode for wide‑format printing or PDF export with Aspose.Cells for .NET.
class SetWorksheetOrientation
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Set the page orientation to Landscape for better wide data presentation
        worksheet.PageSetup.Orientation = PageOrientationType.Landscape;

        // (Optional) Add some sample data to illustrate the effect
        worksheet.Cells["A1"].PutValue("Landscape Orientation Demo");

        // Save the workbook to a file
        workbook.Save("LandscapeOrientation.xlsx", SaveFormat.Xlsx);
    }
}
