// Title: Get Worksheet Paper Height with Aspose.Cells for .NET (C#)
// Description: This example creates an in‑memory Workbook, accesses the active Worksheet, reads the PageSetup.PaperHeight value (in inches), prints it to the console, changes the PaperSize to Letter, reads the updated height, and saves the file. It demonstrates how to query and react to paper dimensions programmatically.
// Keywords: Aspose.Cells PaperHeight | C# PageSetup | worksheet paper size | retrieve default paper height | change PaperSize Aspose.Cells | console output Aspose.Cells
// Common Searches: Aspose.Cells get worksheet paper height | PageSetup PaperHeight C# example | how to read default paper size height Aspose.Cells | C# retrieve paper height after setting PaperSize | print paper dimensions to console with Aspose.Cells
// Developer Intent: Read the PaperHeight property of the active worksheet’s PageSetup and display the value.
// Use Cases: Log the current paper height before sending a worksheet to a printer or PDF converter. | Adjust layout calculations based on the worksheet’s printable area when generating custom reports. | Validate that a workbook conforms to a required paper size by comparing its PaperHeight value.
// AI Prompts: Generate C# code using Aspose.Cells that reads the active worksheet’s PaperHeight and writes it to the console. | Show how to set the worksheet PaperSize to A4, then output the new PaperHeight in centimeters. | Create a reusable method that returns the PaperHeight of any worksheet in both inches and centimeters.

using System;
using Aspose.Cells;

// This example creates an in‑memory Workbook, accesses the active Worksheet, reads the PageSetup.PaperHeight value (in inches), prints it to the console, changes the PaperSize to Letter, reads the updated height, and saves the file. It demonstrates how to query and react to paper dimensions programmatically.
class Program
{
    static void Main()
    {
        // Create a new workbook (in-memory)
        Workbook workbook = new Workbook();

        // Access the first worksheet (active worksheet)
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the PageSetup object for the worksheet
        PageSetup pageSetup = worksheet.PageSetup;

        // Retrieve the paper height (in inches) from the PageSetup
        double paperHeight = pageSetup.PaperHeight;

        // Display the default paper height
        Console.WriteLine("Default Paper Height (inches): " + paperHeight);

        // Change the paper size to Letter to see the updated height
        pageSetup.PaperSize = PaperSizeType.PaperLetter;

        // Retrieve and display the new paper height after changing the size
        Console.WriteLine("Paper Height after setting Letter size (inches): " + pageSetup.PaperHeight);

        // Save the workbook (demonstrates lifecycle usage)
        workbook.Save("PaperHeightDemo.xlsx");
    }
}
