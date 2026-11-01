// Title: Export each worksheet's paper size and dimensions to a CSV file with Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens a workbook, reads the PageSetup PaperSize, PaperWidth, and PaperHeight for every worksheet, and writes the results to a CSV file. | Extend the example to also capture PageSetup orientation and margin values and include them in the CSV export. | Write a reusable method that returns a list of objects containing worksheet name, paper size enum, width, and height, then serialize that list to CSV.
// Common Searches: Aspose.Cells how to export worksheet page setup dimensions to CSV in C# | C# retrieve paper width and height for each sheet using Aspose.Cells | Export Excel sheet paper size and orientation to a CSV file programmatically | Get page setup properties of all worksheets with Aspose.Cells .NET | Write worksheet layout information (paper size, dimensions) to CSV using Aspose.Cells
// Tags: export worksheet paper dimensions csv | Aspose.Cells page setup extraction | C# iterate worksheets page setup | retrieve paper size width height inches | worksheet layout metrics to csv

using Aspose.Cells;
using System;
using System.IO;

// The program loads a workbook, iterates through all worksheets, reads each sheet's PageSetup paper size, width, and height (in inches), and writes the worksheet name along with these dimensions to a CSV file named paper_dimensions.csv.
class ExportPaperDimensions
{
    static void Main()
    {
        // Load the existing workbook (load rule)
        Workbook workbook = new Workbook("input.xlsx");

        // Create and open the CSV file for writing
        using (StreamWriter writer = new StreamWriter("paper_dimensions.csv"))
        {
            // Write CSV header
            writer.WriteLine("Worksheet,PaperSize,PaperWidthInches,PaperHeightInches");

            // Iterate through each worksheet to gather paper dimensions
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the page setup of the worksheet
                PageSetup setup = sheet.PageSetup;

                // Retrieve paper size as a string
                string paperSize = setup.PaperSize.ToString();

                // Retrieve paper width and height in inches
                double width = setup.PaperWidth;
                double height = setup.PaperHeight;

                // Write the collected data as a CSV line
                writer.WriteLine($"{sheet.Name},{paperSize},{width},{height}");
            }
        }

        // No workbook saving required for this export operation
    }
}
