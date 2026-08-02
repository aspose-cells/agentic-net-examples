// Title: C# Console App – List Worksheet Paper Width & Height (inches) Using Aspose.Cells
// Description: A simple C# console program that loads an Excel workbook, iterates through all worksheets, reads each sheet's PageSetup.PaperWidth and PaperHeight (in inches), and writes the worksheet name with its paper dimensions to the console.
// Keywords: Aspose.Cells C# console example | read worksheet paper size | PageSetup PaperWidth PaperHeight | Excel sheet dimensions inches | list worksheet page setup | Aspose.Cells get paper dimensions
// Common Searches: Aspose.Cells get worksheet paper width and height | C# console program to display Excel sheet page size | How to read page setup dimensions with Aspose.Cells | Print Excel sheet paper size in inches using .NET
// Developer Intent: Load an Excel workbook and output each worksheet's paper width and height.
// Use Cases: Verify printing layout before converting to PDF or XPS. | Ensure all sheets conform to a corporate paper‑size standard. | Generate a log of worksheet page‑setup settings for documentation or QA.
// AI Prompts: Create a method that returns a dictionary of worksheet names mapped to their paper width and height values. | Show how to convert the retrieved dimensions from inches to centimeters within the same console app. | Add comprehensive error handling for missing files, unsupported page‑setup properties, and permission issues.

using System;
using Aspose.Cells;

namespace PaperDimensionsApp
{
    // A simple C# console program that loads an Excel workbook, iterates through all worksheets, reads each sheet's PageSetup.PaperWidth and PaperHeight (in inches), and writes the worksheet name with its paper dimensions to the console.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook file (modify as needed)
            string filePath = "input.xlsx";

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(filePath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the page setup of the current worksheet
                PageSetup pageSetup = sheet.PageSetup;

                // Retrieve paper dimensions in inches
                double paperWidth = pageSetup.PaperWidth;
                double paperHeight = pageSetup.PaperHeight;

                // Output the worksheet name and its paper dimensions
                Console.WriteLine($"Worksheet: {sheet.Name}");
                Console.WriteLine($"Paper Width: {paperWidth} inches");
                Console.WriteLine($"Paper Height: {paperHeight} inches");
                Console.WriteLine();
            }
        }
    }
}
