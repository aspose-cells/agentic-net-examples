// Title: Aspose.Cells C# – Disable Row/Column Headers for a Clean Worksheet View
// Description: This example creates a workbook, accesses the first sheet, sets the IsRowColumnHeadersVisible property to false, saves the file as NoHeaders.xlsx, reloads it, and writes the visibility status to the console, demonstrating how to present an Excel sheet without the default row and column labels.
// Keywords: Aspose.Cells C# hide headers | IsRowColumnHeadersVisible property | remove Excel row labels .NET | suppress column headings Aspose | clean worksheet view | Aspose.Cells workbook without grid labels
// Common Searches: C# Aspose.Cells hide row labels | How to turn off column headings in Excel using Aspose | IsRowColumnHeadersVisible false example | Export Excel without row/column headers .NET | Aspose.Cells hide sheet headers programmatically
// Developer Intent: Suppress the display of row and column identifiers in the worksheet UI.
// Use Cases: Produce a printable report where grid labels would distract the reader. | Generate a data file for a third‑party system that expects only cell values. | Embed a workbook in a web viewer that requires a minimalist appearance.
// AI Prompts: Write C# code that sets IsRowColumnHeadersVisible to false for a specific worksheet using Aspose.Cells. | Show how to save a workbook after hiding headers and then verify the setting by reloading it. | Explain how to apply the header‑visibility setting to every sheet in a workbook with Aspose.Cells .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This example creates a workbook, accesses the first sheet, sets the IsRowColumnHeadersVisible property to false, saves the file as NoHeaders.xlsx, reloads it, and writes the visibility status to the console, demonstrating how to present an Excel sheet without the default row and column labels.
    class DisableRowColumnHeaders
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Hide row and column headers in the worksheet view
            worksheet.IsRowColumnHeadersVisible = false;

            // Save the workbook
            workbook.Save("NoHeaders.xlsx");

            // Load the saved workbook to verify the setting
            Workbook loadedWorkbook = new Workbook("NoHeaders.xlsx");
            Worksheet loadedWorksheet = loadedWorkbook.Worksheets[0];
            Console.WriteLine("Row and Column Headers Visible: " + loadedWorksheet.IsRowColumnHeadersVisible);
        }
    }
}
