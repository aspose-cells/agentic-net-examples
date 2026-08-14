// Title: C# – Retrieve each worksheet’s paper width and height with Aspose.Cells
// Description: Load a workbook, iterate through all worksheets, read the PageSetup.PaperWidth and PageSetup.PaperHeight properties, store the dimensions in a Dictionary keyed by sheet name, and display the results.
// Keywords: Aspose.Cells C# worksheet paper size | PageSetup PaperWidth | PageSetup PaperHeight | iterate worksheets Aspose.Cells | dictionary of sheet dimensions | .NET get page dimensions | read worksheet page setup | collect sheet size Aspose
// Common Searches: Aspose.Cells get paper size for each sheet | C# loop worksheets and read PageSetup dimensions | How to store worksheet paper width and height in a dictionary | Retrieve page setup size of all worksheets Aspose.Cells | C# Aspose.Cells print sheet dimensions
// Developer Intent: Extract and keep the paper width and height of every worksheet in a workbook.
// Use Cases: Validate page layout before exporting to PDF or printing. | Programmatically adjust scaling or margins based on sheet size. | Audit a workbook to ensure all sheets share a uniform page format.
// AI Prompts: Generate a C# method that returns a Dictionary<string, (double Width, double Height)> with the paper size of each worksheet in an Aspose.Cells workbook. | Provide code to change any worksheet whose width is under 8 inches to A4 paper size using Aspose.Cells. | Explain how to export the collected sheet dimensions to a CSV file with columns: SheetName, WidthInches, HeightInches.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Load a workbook, iterate through all worksheets, read the PageSetup.PaperWidth and PageSetup.PaperHeight properties, store the dimensions in a Dictionary keyed by sheet name, and display the results.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Dictionary to store each sheet's paper width and height (in inches)
        var paperSizes = new Dictionary<string, (double Width, double Height)>();

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Access the PageSetup of the current worksheet
            PageSetup pageSetup = sheet.PageSetup;

            // Get paper dimensions
            double width = pageSetup.PaperWidth;
            double height = pageSetup.PaperHeight;

            // Store dimensions using the worksheet name as the key
            paperSizes[sheet.Name] = (width, height);
        }

        // Example output of collected paper sizes
        foreach (var entry in paperSizes)
        {
            Console.WriteLine($"Sheet: {entry.Key}, Width: {entry.Value.Width} in, Height: {entry.Value.Height} in");
        }

        // If you made any changes and need to save the workbook, uncomment the line below
        // workbook.Save("output.xlsx");
    }
}
