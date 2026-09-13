// Title: Get the active worksheet's paper height with Aspose.Cells for .NET and output it to the console
// AI Prompts: Write a C# console program that loads a workbook using Aspose.Cells, accesses the active worksheet's PageSetup, and prints the PaperHeight value. | Show how to retrieve the PaperHeight property from the active sheet's PageSetup in Aspose.Cells and display the result in the console. | Generate minimal Aspose.Cells code that reads the current paper height setting of the workbook's active worksheet and writes it to standard output.
// Common Searches: asp.net aspose.cells retrieve active worksheet paper height | c# get worksheet page setup paper height using Aspose.Cells | how to print worksheet PaperHeight to console with Aspose.Cells .NET
// Tags: Aspose.Cells active worksheet PaperHeight | C# PageSetup retrieve paper height | console output worksheet paper dimensions | Aspose.Cells get page setup size | read active sheet paper height .NET

using System;
using Aspose.Cells;

// Demonstrates how to obtain the PaperHeight from the active worksheet's PageSetup using Aspose.Cells for .NET and print the value to the console.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one if needed)
        Workbook workbook = new Workbook();

        // Get the active worksheet
        Worksheet activeSheet = workbook.Worksheets[workbook.Worksheets.ActiveSheetIndex];

        // Retrieve the paper height from the worksheet's PageSetup
        double paperHeight = activeSheet.PageSetup.PaperHeight;

        // Display the paper height in the console
        Console.WriteLine($"Paper Height: {paperHeight}");
    }
}
