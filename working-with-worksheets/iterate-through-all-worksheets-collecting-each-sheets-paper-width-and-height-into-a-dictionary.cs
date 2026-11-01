// Title: Retrieve each worksheet’s paper width and height with Aspose.Cells for .NET and store them in a dictionary
// AI Prompts: Write C# code using Aspose.Cells that iterates over every worksheet in a workbook, reads the PageSetup.PaperWidth and PageSetup.PaperHeight properties, and adds the values to a Dictionary<string, (double Width, double Height)> keyed by the worksheet name. | Modify the previous snippet to convert the retrieved paper dimensions from points to millimeters before inserting them into the dictionary. | Enhance the solution with try‑catch blocks around the worksheet loop and export the final dictionary to a CSV file containing columns: SheetName, WidthMm, HeightMm.
// Common Searches: Aspose.Cells get page setup dimensions for each sheet in C# | C# code to read worksheet paper size (width height) from an Excel file using Aspose | How to store Excel sheet paper dimensions in a .NET dictionary with Aspose.Cells | Convert Aspose.Cells PaperWidth from points to millimeters example | Export worksheet paper size information to CSV using Aspose.Cells
// Tags: aspnet worksheet pagesetup paperwidth retrieval | dictionary of sheet paper dimensions c# | paper size unit conversion aspose.cells | csv output of worksheet dimensions aspose.cells | iterate worksheets aspose.cells page setup

using Aspose.Cells;
using System;
using System.Collections.Generic;

// Loads an Excel workbook with Aspose.Cells, iterates each worksheet, reads its PageSetup PaperWidth and PaperHeight (in points), stores the dimensions in a dictionary keyed by worksheet name, and can convert units or write the results to a CSV file.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Dictionary to store each sheet's paper width and height (in points)
        var sheetPaperSizes = new Dictionary<string, (double Width, double Height)>();

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Access the page setup of the current worksheet
            PageSetup pageSetup = sheet.PageSetup;

            // Retrieve paper width and height (values are in points)
            double width = pageSetup.PaperWidth;
            double height = pageSetup.PaperHeight;

            // Store the dimensions using the sheet name as the key
            sheetPaperSizes[sheet.Name] = (width, height);
        }

        // Example output of the collected paper sizes
        foreach (var kvp in sheetPaperSizes)
        {
            Console.WriteLine($"Sheet: {kvp.Key}, Width: {kvp.Value.Width} pt, Height: {kvp.Value.Height} pt");
        }
    }
}
