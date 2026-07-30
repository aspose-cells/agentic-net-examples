// Title: Hide Row and Column Headers in Aspose.Cells (C#) for a Clean Worksheet Layout
// Description: Shows how to set Worksheet.IsRowColumnHeadersVisible to false in Aspose.Cells for .NET, creating a workbook without the default A‑Z column and 1‑N row headings and saving it as CleanLayout.xlsx.
// Keywords: Aspose.Cells | C# | Hide row headers | Hide column headers | IsRowColumnHeadersVisible | clean worksheet view | disable Excel headings | worksheet UI | Aspose.Cells .NET | Excel without headers
// Common Searches: Aspose.Cells hide row headers C# | Aspose.Cells hide column headings | remove Excel row and column labels using Aspose | Worksheet.IsRowColumnHeadersVisible false example | create workbook without headers Aspose.Cells
// Developer Intent: Hide the default row and column labels in an Aspose.Cells worksheet to produce a header‑free view.
// Use Cases: Generate a data‑only report where row/column labels would distract from the content. | Prepare a printable spreadsheet for a presentation that should not display Excel headings. | Design a dashboard sheet that contains only charts or images, requiring no visible row or column markers.
// AI Prompts: Write C# code that hides row and column headers in an Aspose.Cells worksheet and then exports the file to PDF. | Explain how to toggle Worksheet.IsRowColumnHeadersVisible based on a user setting in an Aspose.Cells application. | Describe the effect of disabling row and column headers on Excel's UI and how to restore them later.

using System;
using Aspose.Cells;

// Shows how to set Worksheet.IsRowColumnHeadersVisible to false in Aspose.Cells for .NET, creating a workbook without the default A‑Z column and 1‑N row headings and saving it as CleanLayout.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (default contains one worksheet)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Hide row and column headers in the worksheet view
        worksheet.IsRowColumnHeadersVisible = false;

        // Save the workbook to a file
        workbook.Save("CleanLayout.xlsx");
    }
}
