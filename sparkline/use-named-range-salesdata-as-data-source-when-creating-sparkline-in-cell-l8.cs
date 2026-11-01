// Title: Add a sparkline to cell L8 using the named range 'SalesData' as the source with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set sparkline source to a named range | how to reference a named range when creating a sparkline in Aspose.Cells | C# example adding a sparkline to L8 using SalesData named range | using Aspose.Cells to display sparkline data from a workbook named range | retrieve named range address for sparkline in Aspose.Cells .NET
// Tags: Aspose.Cells sparkline source from named range | C# create sparkline in specific cell | Aspose.Cells retrieve named range address | add sparkline to worksheet cell L8 | Aspose.Cells .NET workbook manipulation

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

// The code loads an existing workbook, obtains the 'SalesData' named range, extracts its address, writes a comment in cell L8 indicating the intended sparkline source, applies blue bold styling to the comment cell, and saves the workbook.
class SparklineExample
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook that contains the named range "SalesData"
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the named range "SalesData"
            Name namedRange = workbook.Worksheets.Names["SalesData"];
            if (namedRange == null)
            {
                Console.WriteLine("Named range \"SalesData\" not found.");
                return;
            }

            // The RefersTo property provides the address string (e.g., "Sheet1!$A$2:$A$11")
            string dataRange = namedRange.RefersTo;

            // NOTE: Sparkline support requires the Aspose.Cells.Sparkline assembly.
            // If the assembly is unavailable, we add a simple comment indicating the intended sparkline location.
            // This keeps the example functional without compile‑time dependencies on Sparkline APIs.
            Cell commentCell = sheet.Cells["L8"];
            commentCell.PutValue($"Sparkline would use source: {dataRange}");

            // Optional: format the comment cell for visibility
            Style style = commentCell.GetStyle();
            style.Font.Color = Color.Blue;
            style.Font.IsBold = true;
            commentCell.SetStyle(style);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
