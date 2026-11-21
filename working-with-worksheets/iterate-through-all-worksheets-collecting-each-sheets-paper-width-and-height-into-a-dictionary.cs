// Title: Retrieve each worksheet's paper width and height and store them in a C# dictionary with Aspose.Cells
// Common Searches: Aspose.Cells get paper size of every worksheet in a workbook C# | C# loop through worksheets and read PageSetup PaperWidth and PaperHeight | store worksheet page dimensions in a dictionary using Aspose.Cells | extract sheet paper dimensions from an Excel file with Aspose.Cells .NET | read worksheet print layout size via Aspose.Cells PageSetup properties
// Tags: Aspose.Cells worksheet page setup dimensions | C# dictionary of sheet paper sizes | iterate workbook worksheets Aspose.Cells | read PaperWidth PaperHeight Aspose.Cells | collect worksheet print dimensions .NET

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The example loads an Excel workbook, creates a Dictionary keyed by worksheet name, iterates through all worksheets, accesses each sheet's PageSetup to obtain PaperWidth and PaperHeight (in points), stores these values as a tuple in the dictionary, and then outputs the collected dimensions.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Dictionary to hold sheet name as key and its paper width & height as value
        var sheetPaperSizes = new Dictionary<string, (double Width, double Height)>();

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Access the PageSetup of the current worksheet
            PageSetup pageSetup = sheet.PageSetup;

            // Retrieve paper width and height (in points)
            double width = pageSetup.PaperWidth;
            double height = pageSetup.PaperHeight;

            // Store the values in the dictionary using the sheet name as the key
            sheetPaperSizes[sheet.Name] = (width, height);
        }

        // Example usage: print collected paper sizes
        foreach (var kvp in sheetPaperSizes)
        {
            Console.WriteLine($"Sheet: {kvp.Key}, Width: {kvp.Value.Width}, Height: {kvp.Value.Height}");
        }

        // (Optional) Save the workbook if any modifications were made
        // workbook.Save("output.xlsx");
    }
}
