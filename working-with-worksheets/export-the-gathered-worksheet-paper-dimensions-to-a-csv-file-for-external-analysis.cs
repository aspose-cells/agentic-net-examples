// Title: Export worksheet paper size, width, and height to a CSV file using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export each worksheet's paper size and dimensions to CSV | How to retrieve page setup paper width and height for all sheets with Aspose.Cells | Save Excel worksheet page setup properties as CSV using Aspose.Cells .NET | Aspose.Cells iterate worksheets and write paper dimensions to CSV file | Export Excel sheet paper size enum and dimensions to CSV in C#
// Tags: Aspose.Cells worksheet page setup CSV export | extract paper size enum Aspose.Cells | write paper width and height to CSV C# | iterate workbook worksheets Aspose.Cells | page setup dimensions export .NET

using Aspose.Cells;
using System;
using System.IO;

// The C# program loads an Excel workbook with Aspose.Cells, loops through every worksheet, reads the PageSetup properties (PaperSize, PaperWidth, PaperHeight), and writes the worksheet name together with these dimensions to a CSV file named "paper_dimensions.csv" for external analysis.
class ExportPaperDimensions
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Create a CSV file to store the paper dimensions
        using (StreamWriter writer = new StreamWriter("paper_dimensions.csv"))
        {
            // Write CSV header
            writer.WriteLine("Worksheet,PaperSize,PaperWidth,PaperHeight");

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the page setup of the current worksheet
                PageSetup pageSetup = sheet.PageSetup;

                // Retrieve paper size (enum name) and dimensions (points)
                string paperSize = pageSetup.PaperSize.ToString();
                double paperWidth = pageSetup.PaperWidth;   // width in points (1 point = 1/72 inch)
                double paperHeight = pageSetup.PaperHeight; // height in points

                // Write the data as a CSV line
                writer.WriteLine($"{sheet.Name},{paperSize},{paperWidth},{paperHeight}");
            }
        }

        // No workbook saving needed; the CSV file contains the exported data.
    }
}
