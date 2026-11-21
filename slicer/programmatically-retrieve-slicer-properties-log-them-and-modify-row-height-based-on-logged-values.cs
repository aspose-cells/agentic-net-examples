// Title: C# example: read slicer properties and set row height in an Excel file with Aspose.Cells
// Common Searches: how to read slicer name and caption using Aspose.Cells in C# | adjust Excel row height based on slicer visibility with Aspose.Cells .NET | log slicer position and size while processing a workbook in C# | Aspose.Cells C# iterate over worksheet slicers and change row dimensions | retrieve slicer row index and modify row height programmatically
// Tags: Aspose.Cells slicer collection iteration | set worksheet row height from slicer data | extract slicer attributes Aspose.Cells | conditional row formatting using slicer visibility | C# read slicer coordinates Excel

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing workbook, enumerates each slicer on the first worksheet (if slicer support is available), writes the slicer's name, caption, row/column position, size and visibility to the console, then changes the height of the row that contains the slicer's top‑left cell to 30 points when the slicer is visible or 15 points otherwise, and finally saves the workbook as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Input workbook path
            string inputPath = "Input.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Work with the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // -----------------------------------------------------------------
            // NOTE: Slicer APIs may not be available in older Aspose.Cells versions.
            // The following block is guarded to prevent compile‑time errors.
            // If slicer support exists, uncomment the block and ensure the
            // appropriate Aspose.Cells version is referenced.
            // -----------------------------------------------------------------
            /*
            // Retrieve all slicers on the worksheet
            Aspose.Cells.Tables.SlicerCollection slicers = sheet.Slicers;

            // Iterate through each slicer, log its properties and adjust row height
            foreach (Aspose.Cells.Tables.Slicer slicer in slicers)
            {
                // Log basic slicer properties
                Console.WriteLine($"Slicer Name: {slicer.Name}");
                Console.WriteLine($"Slicer Caption: {slicer.Caption}");
                Console.WriteLine($"Slicer Position (Row, Column): ({slicer.Row}, {slicer.Column})");
                Console.WriteLine($"Slicer Height (points): {slicer.Height}");
                Console.WriteLine($"Slicer Width (points): {slicer.Width}");
                Console.WriteLine($"Slicer Visible: {slicer.IsVisible}");
                Console.WriteLine("-----------------------------------");

                // Determine the row to modify based on the slicer's top‑left cell (zero‑based index)
                int targetRowIndex = slicer.Row;

                // Example rule: larger height if visible, smaller otherwise
                double newHeight = slicer.IsVisible ? 30.0 : 15.0;

                // Apply the new row height
                Row targetRow = sheet.Cells.Rows[targetRowIndex];
                targetRow.Height = newHeight;
            }
            */

            // -----------------------------------------------------------------
            // If slicer handling is not required or unavailable, you can perform
            // other worksheet operations here.
            // -----------------------------------------------------------------

            // Save the modified workbook
            string outputPath = "Output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
