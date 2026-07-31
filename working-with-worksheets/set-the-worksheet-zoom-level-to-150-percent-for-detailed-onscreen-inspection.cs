// Title: Set Excel Worksheet Zoom to 150% with Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, accesses the first worksheet, sets its Zoom property to 150 % for a closer on‑screen view, writes the zoom value to the console, and saves the file as WorksheetZoom150.xlsx.
// Keywords: Aspose.Cells worksheet zoom | C# set Excel zoom | Aspose.Cells Zoom property | 150 percent view Aspose | default worksheet zoom .NET | Excel workbook zoom example | Aspose.Cells .NET US developers | Aspose.Cells tutorial India
// Common Searches: Aspose.Cells set worksheet zoom 150 | C# change Excel worksheet view zoom factor | How to set default zoom in an Excel file using Aspose.Cells | Increase worksheet zoom programmatically .NET | Aspose.Cells zoom property example C#
// Developer Intent: Apply a 150 % view zoom to a worksheet before saving the workbook.
// Use Cases: Prepare a workbook that opens already zoomed in for detailed data inspection. | Generate a preview file where the default view is enlarged for better readability on screens. | Create a template that forces end‑users to see the worksheet at a specific zoom level when opened.
// AI Prompts: Show C# code to set the zoom level of a specific worksheet to 150 % using Aspose.Cells for .NET. | Provide an example that changes the Zoom property for all worksheets in a workbook to 150 % with Aspose.Cells. | Explain the limits and behavior of the Worksheet.Zoom property in Aspose.Cells. | How can I make an Excel file open with a predefined zoom level using Aspose.Cells in a .NET application?

using System;
using Aspose.Cells;

namespace AsposeCellsZoomDemo
{
    // Creates a new workbook, accesses the first worksheet, sets its Zoom property to 150 % for a closer on‑screen view, writes the zoom value to the console, and saves the file as WorksheetZoom150.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (default contains one worksheet)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set the zoom factor to 150%
            worksheet.Zoom = 150;

            // Optional: display the current zoom factor in console
            Console.WriteLine("Worksheet zoom set to: " + worksheet.Zoom + "%");

            // Save the workbook to a file
            workbook.Save("WorksheetZoom150.xlsx");
        }
    }
}
