// Title: Aspose.Cells C# – Retrieve Worksheet Paper Height (PageSetup) and Output to Console
// Description: Demonstrates how to access the active worksheet's PageSetup in Aspose.Cells for .NET, read the PaperHeight property (in inches), write the value to the console, and save the workbook.
// Keywords: Aspose.Cells PaperHeight | C# PageSetup PaperHeight | read worksheet paper size Aspose | console output paper height | Aspose.Cells page setup example
// Common Searches: Aspose.Cells get paper height C# | PageSetup PaperHeight property example | how to read worksheet paper size with Aspose | display paper dimensions in console Aspose.Cells | retrieve active worksheet page setup dimensions
// Developer Intent: Read the PaperHeight value of the active worksheet's PageSetup and display it in the console.
// Use Cases: Log paper height before printing to verify page size. | Compare paper dimensions across multiple worksheets for consistent report layout. | Adjust scaling or layout logic based on the retrieved paper height.
// AI Prompts: Generate C# code that reads both PaperWidth and PaperHeight from a worksheet's PageSetup using Aspose.Cells and formats the output as JSON. | Show how to compare PaperHeight of two worksheets and set a custom page size when they differ, using Aspose.Cells for .NET. | Explain how to convert the PaperHeight value from inches to millimeters and display both units in the console.

using System;
using Aspose.Cells;

// Demonstrates how to access the active worksheet's PageSetup in Aspose.Cells for .NET, read the PaperHeight property (in inches), write the value to the console, and save the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook (in-memory)
        Workbook workbook = new Workbook();

        // Access the first (active) worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the PageSetup object for the worksheet
        PageSetup pageSetup = worksheet.PageSetup;

        // Retrieve the paper height in inches
        double paperHeight = pageSetup.PaperHeight;

        // Display the paper height in the console
        Console.WriteLine("Paper Height (inches): " + paperHeight);

        // Save the workbook (demonstrates lifecycle handling)
        workbook.Save("PaperHeightDemo.xlsx");
    }
}
