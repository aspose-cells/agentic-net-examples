// Title: Get the worksheet paper height with Aspose.Cells for .NET and print it to the console
// Common Searches: Aspose.Cells C# get worksheet page setup paper height | how to read paper height of active sheet using Aspose.Cells | display worksheet paper height in console with Aspose.Cells | retrieve page setup dimensions from workbook in .NET
// Tags: Aspose.Cells page setup paper height property | C# active worksheet paper size retrieval | output worksheet dimensions to console | paper height points Aspose.Cells API

using System;
using Aspose.Cells;

// The example creates a Workbook, accesses the active Worksheet, reads the PaperHeight value (in points) from the worksheet's PageSetup, and writes the height to the console.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one if needed)
        Workbook workbook = new Workbook();

        // Get the active worksheet
        Worksheet activeSheet = workbook.Worksheets[workbook.Worksheets.ActiveSheetIndex];

        // Retrieve the paper height (in points) from the worksheet's page setup
        double paperHeight = activeSheet.PageSetup.PaperHeight;

        // Display the paper height in the console
        Console.WriteLine($"Paper Height: {paperHeight} points");
    }
}
