// Title: Set Worksheet Zoom to 150% with Aspose.Cells for .NET (C#)
// Description: This C# example demonstrates how to create a workbook with Aspose.Cells, access the first worksheet, set its Zoom property to 150 %, output the value to the console, and save the file as WorksheetZoom150.xlsx.
// Keywords: Aspose.Cells | C# worksheet zoom | set Excel zoom programmatically | Worksheet.Zoom property | 150 percent zoom | Aspose.Cells .NET example | adjust worksheet view | Excel sheet zoom level
// Common Searches: Aspose.Cells set worksheet zoom C# | how to change Excel sheet zoom to 150% using .NET | Worksheet.Zoom property example | increase Excel zoom programmatically C# | set default zoom in Aspose.Cells workbook
// Developer Intent: Apply a 150 % zoom level to a worksheet using Aspose.Cells for .NET.
// Use Cases: Generate reports that open with a larger view for detailed inspection | Debug layout issues by programmatically enlarging the sheet view | Create templates that open at a predefined zoom for consistent user experience | Automate zoom adjustments based on user preferences before exporting
// AI Prompts: Show code to set different zoom percentages for each worksheet in a workbook with Aspose.Cells. | Provide a C# snippet that reads the current Zoom value, compares it to a threshold, and modifies it accordingly. | Explain how to reset a worksheet's zoom to the default 100 % after custom changes using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsZoomDemo
{
    // This C# example demonstrates how to create a workbook with Aspose.Cells, access the first worksheet, set its Zoom property to 150 %, output the value to the console, and save the file as WorksheetZoom150.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set the zoom factor to 150%
            worksheet.Zoom = 150;

            // Optionally display the current zoom factor
            Console.WriteLine("Worksheet zoom set to: " + worksheet.Zoom + "%");

            // Save the workbook to a file
            workbook.Save("WorksheetZoom150.xlsx");
        }
    }
}
