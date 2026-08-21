// Title: Set Worksheet Zoom to 150% with Aspose.Cells for .NET
// Description: Demonstrates how to use the Worksheet.Zoom property in Aspose.Cells for .NET to display the first sheet at 150 % zoom, verify the setting, and save the workbook.
// Keywords: Aspose.Cells Worksheet.Zoom | C# set Excel sheet zoom | programmatically change Excel zoom .NET | Aspose.Cells zoom percentage | adjust worksheet view scale
// Common Searches: Aspose.Cells set worksheet zoom .NET | C# change Excel sheet zoom programmatically | How to increase worksheet zoom to 150% using Aspose | Set Excel view zoom with Aspose.Cells | Worksheet.Zoom property example
// Developer Intent: Apply a 150 % zoom to a worksheet so the file opens with a larger view for detailed inspection.
// Use Cases: Generate a report that opens with the first sheet enlarged for easier reading. | Create a template where every worksheet defaults to 150 % zoom for fine‑grained data analysis. | Prepare a presentation workbook that automatically shows charts at a higher magnification.
// AI Prompts: Show C# code that sets the Zoom property of a specific worksheet to 150 % with Aspose.Cells and prints the value. | Provide a snippet to loop through all worksheets in a workbook and set each Zoom to 150 % using Aspose.Cells. | Explain how the Worksheet.Zoom setting influences the initial view in Excel, Google Sheets, and other spreadsheet viewers.

using System;
using Aspose.Cells;

// Demonstrates how to use the Worksheet.Zoom property in Aspose.Cells for .NET to display the first sheet at 150 % zoom, verify the setting, and save the workbook.
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

        // Output the current zoom factor for verification
        Console.WriteLine($"Current worksheet zoom: {worksheet.Zoom}%");

        // Save the workbook
        workbook.Save("ZoomedWorksheet.xlsx");
    }
}
