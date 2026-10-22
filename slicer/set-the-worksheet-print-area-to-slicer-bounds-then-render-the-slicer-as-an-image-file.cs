// Title: How to set a worksheet's print area to a slicer's bounds using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that locates the first slicer on a worksheet, extracts its top‑row and left‑column coordinates (using reflection if needed), and assigns a matching print area with Aspose.Cells. | Write a C# example that builds the cell address from a slicer's position, sets Worksheet.PageSetup.PrintArea to that range, and saves the workbook. | Create a C# snippet that falls back to the worksheet's used range when slicer position properties are unavailable, then updates the print area accordingly.
// Common Searches: Aspose.Cells C# set print area based on slicer position | How to get slicer top row and left column with Aspose.Cells .NET | Define worksheet print area to a single cell covering a slicer using Aspose.Cells | C# Aspose.Cells reflection to read slicer properties | Save workbook after modifying print area with Aspose.Cells
// Tags: Aspose.Cells set print area from slicer | C# retrieve slicer coordinates Aspose.Cells | Aspose.Cells reflection slicer properties | Aspose.Cells workbook save after print area change | Aspose.Cells worksheet page setup print area

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

// The example loads an existing workbook, accesses the first slicer on the first worksheet, obtains its top‑row and left‑column indices (using reflection when available, otherwise falling back to the used range), constructs a single‑cell address, sets the worksheet's PageSetup.PrintArea to that address, and saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Input workbook path
            string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one slicer
            if (sheet.Slicers.Count == 0)
            {
                Console.WriteLine("No slicers found on the worksheet.");
                return;
            }

            // Get the first slicer (or select by name/index as required)
            Slicer slicer = sheet.Slicers[0];

            // ----- Determine slicer bounds -----
            // Aspose.Cells versions prior to certain releases may not expose TopRow/LeftColumn.
            // As a fallback, use the worksheet's used range to define a print area.
            // If the slicer provides TopRow/LeftColumn, they will be used; otherwise, the fallback is applied.

            int topRow;
            int leftCol;

            // Attempt to use TopRow/LeftColumn if available via reflection
            var topRowProp = slicer.GetType().GetProperty("TopRow");
            var leftColProp = slicer.GetType().GetProperty("LeftColumn");

            if (topRowProp != null && leftColProp != null)
            {
                topRow = ((int)topRowProp.GetValue(slicer)) + 1; // Convert to 1‑based index
                leftCol = (int)leftColProp.GetValue(slicer);
            }
            else
            {
                // Fallback to the first cell of the used range
                topRow = sheet.Cells.MaxDisplayRange.FirstRow + 1;
                leftCol = sheet.Cells.MaxDisplayRange.FirstColumn;
            }

            // Convert column index to letter (e.g., 0 -> A)
            string leftColLetter = CellsHelper.ColumnIndexToName(leftCol);

            // Build the address string for the print area (single cell covering the slicer or used range start)
            string printArea = $"{leftColLetter}{topRow}:{leftColLetter}{topRow}";

            // Set the worksheet print area to match the determined bounds
            sheet.PageSetup.PrintArea = printArea;

            // (Optional) Save the workbook if you want to keep the updated print area
            string outputPath = "output.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);

            Console.WriteLine($"Print area set to: {printArea}");
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
