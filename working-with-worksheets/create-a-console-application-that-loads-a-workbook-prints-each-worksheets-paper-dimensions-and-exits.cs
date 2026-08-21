// Title: C# Console App – List Worksheet Paper Width & Height with Aspose.Cells
// Description: A simple C# console program that loads an Excel workbook using Aspose.Cells for .NET, iterates through all worksheets, and prints each sheet's PageSetup.PaperWidth and PaperHeight (in inches) to the console.
// Keywords: Aspose.Cells | C# | Console application | Workbook | Worksheet | PaperWidth | PaperHeight | PageSetup | Excel print dimensions | Get sheet size .NET
// Common Searches: How to read worksheet paper size with Aspose.Cells C# | Aspose.Cells console example for page dimensions | Retrieve Excel sheet width and height using .NET | List page setup dimensions of each worksheet | C# program to display Excel print layout sizes
// Developer Intent: The developer needs a quick way to open an Excel file and output the printable width and height of every worksheet via the console.
// Use Cases: Verify print layout settings before converting sheets to PDF or XPS. | Audit page setup consistency across multiple worksheets in a workbook. | Log or export worksheet paper dimensions for automated reporting pipelines.
// AI Prompts: Create a method that returns a dictionary of worksheet names mapped to their PaperWidth and PaperHeight values. | Add error handling for missing files, unsupported page setups, and non‑existent worksheets. | Enhance the app to write the collected dimensions to a CSV or JSON file.

using System;
using Aspose.Cells;

namespace PaperDimensionsApp
{
    // A simple C# console program that loads an Excel workbook using Aspose.Cells for .NET, iterates through all worksheets, and prints each sheet's PageSetup.PaperWidth and PaperHeight (in inches) to the console.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook file (modify as needed)
            string filePath = "input.xlsx";

            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Print paper dimensions for each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                PageSetup setup = sheet.PageSetup;
                Console.WriteLine($"Worksheet \"{sheet.Name}\":");
                Console.WriteLine($"  Paper Width  = {setup.PaperWidth} inches");
                Console.WriteLine($"  Paper Height = {setup.PaperHeight} inches");
            }
        }
    }
}
