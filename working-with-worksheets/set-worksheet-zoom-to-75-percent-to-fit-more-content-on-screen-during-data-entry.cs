// Title: C# – Set worksheet zoom to 75% with Aspose.Cells for .NET
// Description: This example creates a new Workbook, accesses the first Worksheet, sets its Zoom property to 75 %, prints the applied value, and saves the file as WorksheetZoom75.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# | worksheet zoom | Zoom property | set Excel view programmatically | 75 percent zoom | Aspose.Cells for .NET example | GitHub Aspose.Cells sample | Excel workbook zoom
// Common Searches: Aspose.Cells set worksheet zoom C# | how to change Excel zoom with Aspose.Cells | C# code to set worksheet view to 75% | Aspose.Cells Zoom property example | programmatically adjust Excel zoom level
// Developer Intent: Set the worksheet's view zoom to 75 % before saving the workbook.
// Use Cases: Provide a default 75 % zoom for data‑entry worksheets to improve on‑screen readability. | Standardize the initial view of generated reports across all users. | Apply user‑defined zoom levels dynamically when exporting workbooks.
// AI Prompts: Write C# code that uses Aspose.Cells to set a worksheet's zoom to a variable percentage and save the workbook. | Explain the range and behavior of the Zoom property in Aspose.Cells, including how to read the current zoom factor. | Show how to assign different zoom percentages to multiple worksheets within the same workbook using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsZoomDemo
{
    // This example creates a new Workbook, accesses the first Worksheet, sets its Zoom property to 75 %, prints the applied value, and saves the file as WorksheetZoom75.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set the worksheet zoom to 75%
            worksheet.Zoom = 75;

            // Optional: display the current zoom factor
            Console.WriteLine("Worksheet zoom set to: " + worksheet.Zoom + "%");

            // Save the workbook (lifecycle: save)
            workbook.Save("WorksheetZoom75.xlsx");
        }
    }
}
