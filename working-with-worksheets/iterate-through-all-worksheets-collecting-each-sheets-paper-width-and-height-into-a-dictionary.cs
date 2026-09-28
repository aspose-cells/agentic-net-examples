// Title: Extract worksheet paper width and height into a dictionary using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel file with Aspose.Cells, iterates over every worksheet, reads the PageSetup.PaperWidth and PageSetup.PaperHeight properties, and stores them in a Dictionary<string, (double Width, double Height)> keyed by worksheet name. | Provide a C# example that prints each worksheet’s name together with its paper dimensions in points after collecting them into a dictionary using Aspose.Cells.
// Common Searches: Aspose.Cells C# get paper size of each worksheet | How to read PageSetup PaperWidth and PaperHeight for all sheets in a workbook | C# dictionary of worksheet names and print layout dimensions using Aspose.Cells | Iterate worksheets and retrieve print page dimensions with Aspose.Cells .NET
// Tags: Aspose.Cells worksheet page setup size extraction | C# map worksheet name to paper dimensions | collect PageSetup PaperWidth PaperHeight Aspose.Cells | iterate Excel worksheets for print layout metrics | dictionary of worksheet print size .NET

using Aspose.Cells;
using System;
using System.Collections.Generic;

// Loads an Excel workbook with Aspose.Cells, loops through all worksheets, reads each sheet’s PageSetup.PaperWidth and PaperHeight (points), stores the values in a Dictionary keyed by worksheet name, and outputs the dimensions.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Dictionary to store each sheet's paper width and height
        // Key: worksheet name, Value: tuple of (width, height) in points
        Dictionary<string, (double Width, double Height)> sheetPaperSizes = new Dictionary<string, (double, double)>();

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Retrieve paper dimensions from the worksheet's PageSetup
            double paperWidth = sheet.PageSetup.PaperWidth;   // width in points
            double paperHeight = sheet.PageSetup.PaperHeight; // height in points

            // Add the dimensions to the dictionary using the sheet name as the key
            sheetPaperSizes[sheet.Name] = (paperWidth, paperHeight);
        }

        // Example output to verify the collected data
        foreach (var entry in sheetPaperSizes)
        {
            Console.WriteLine($"Sheet: {entry.Key}, Width: {entry.Value.Width} pt, Height: {entry.Value.Height} pt");
        }
    }
}
