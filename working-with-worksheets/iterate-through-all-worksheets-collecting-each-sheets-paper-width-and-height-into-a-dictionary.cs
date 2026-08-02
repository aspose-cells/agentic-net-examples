// Title: Retrieve each worksheet's paper width and height into a dictionary with Aspose.Cells for .NET (C#)
// Description: This C# example creates a workbook, assigns different PaperSize values to three sheets, iterates all worksheets, reads PageSetup.PaperWidth and PageSetup.PaperHeight (in inches), stores the dimensions in a Dictionary keyed by sheet name, prints the results, and saves the file.
// Keywords: Aspose.Cells | C# | .NET | worksheet paper size | PageSetup PaperWidth | PageSetup PaperHeight | dictionary of sheet dimensions | iterate worksheets | retrieve page dimensions | paper size Aspose.Cells | Aspose.Cells sample code
// Common Searches: Aspose.Cells get worksheet paper size C# | How to read PaperWidth and PaperHeight for each sheet in Aspose.Cells | C# dictionary of worksheet page dimensions Aspose.Cells | Iterate all worksheets and collect page setup size Aspose.Cells | Aspose.Cells .NET example for paper dimensions
// Developer Intent: Extract the paper width and height of every worksheet and store them in a dictionary keyed by the worksheet name.
// Use Cases: Validate layout before printing by listing each sheet's page dimensions. | Adjust scaling or margins programmatically based on collected paper sizes. | Export sheet dimension data to external systems for batch printing or reporting. | Create a summary sheet that lists page sizes for documentation.
// AI Prompts: Generate C# code using Aspose.Cells to loop through all worksheets in a workbook and build a Dictionary<string, (double Width, double Height)> containing each sheet's PaperWidth and PaperHeight in inches. | Show how to log each worksheet name with its paper dimensions and then save the workbook using Aspose.Cells. | Provide a concise Aspose.Cells .NET snippet that reads PageSetup.PaperWidth and PaperHeight for every worksheet and returns the results as a JSON object.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsPaperSizeDemo
{
    // This C# example creates a workbook, assigns different PaperSize values to three sheets, iterates all worksheets, reads PageSetup.PaperWidth and PageSetup.PaperHeight (in inches), stores the dimensions in a Dictionary keyed by sheet name, prints the results, and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Example: add a few worksheets with different paper sizes
            workbook.Worksheets[0].Name = "Sheet1";
            workbook.Worksheets[0].PageSetup.PaperSize = PaperSizeType.PaperA4;

            int sheet2Idx = workbook.Worksheets.Add();
            workbook.Worksheets[sheet2Idx].Name = "Sheet2";
            workbook.Worksheets[sheet2Idx].PageSetup.PaperSize = PaperSizeType.PaperLetter;

            int sheet3Idx = workbook.Worksheets.Add();
            workbook.Worksheets[sheet3Idx].Name = "Sheet3";
            workbook.Worksheets[sheet3Idx].PageSetup.PaperSize = PaperSizeType.PaperLegal;

            // Dictionary to hold sheet name -> (PaperWidth, PaperHeight)
            Dictionary<string, (double Width, double Height)> sheetPaperSizes = new Dictionary<string, (double, double)>();

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the PageSetup of the worksheet
                PageSetup pageSetup = sheet.PageSetup;

                // Retrieve width and height (in inches)
                double width = pageSetup.PaperWidth;
                double height = pageSetup.PaperHeight;

                // Store in dictionary using the sheet's name as the key
                sheetPaperSizes[sheet.Name] = (width, height);
            }

            // Output the collected values
            foreach (var kvp in sheetPaperSizes)
            {
                Console.WriteLine($"Worksheet: {kvp.Key}, Paper Width: {kvp.Value.Width} in, Paper Height: {kvp.Value.Height} in");
            }

            // Save the workbook if needed
            workbook.Save("PaperSizeDemo.xlsx");
        }
    }
}
