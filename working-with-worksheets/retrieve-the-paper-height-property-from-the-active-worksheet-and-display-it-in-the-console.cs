// Title: Read the active worksheet's paper height with Aspose.Cells for .NET and output it to the console
// AI Prompts: Generate C# code that accesses the active worksheet's PageSetup and prints the PaperHeight value in points. | Show how to retrieve the worksheet paper height using Aspose.Cells and display it on the console without loading an external file.
// Common Searches: aspocells c# get active worksheet page setup paper height | how to print Excel worksheet paper height using Aspose.Cells .NET | retrieve paper height from worksheet PageSetup Aspose.Cells example | display worksheet page dimensions in console with Aspose.Cells | Aspose.Cells get paper size height of active sheet
// Tags: Aspose.Cells read worksheet PaperHeight | C# PageSetup PaperHeight extraction | output worksheet page dimensions console | active sheet page setup size retrieval | Aspose.Cells paper size height

using Aspose.Cells;
using System;

// The sample creates a Workbook, obtains the active Worksheet, reads its PageSetup.PaperHeight property (in points), and writes the height value to the console.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one if needed)
        Workbook workbook = new Workbook();

        // Get the active worksheet
        Worksheet activeSheet = workbook.Worksheets[workbook.Worksheets.ActiveSheetIndex];

        // Retrieve the paper height (in points) from the worksheet's PageSetup
        double paperHeight = activeSheet.PageSetup.PaperHeight;

        // Display the paper height in the console
        Console.WriteLine($"Paper Height: {paperHeight} points");
    }
}
