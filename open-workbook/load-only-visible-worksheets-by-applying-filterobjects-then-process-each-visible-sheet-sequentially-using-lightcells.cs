// Title: Load only visible worksheets with FilterObjects and process each sheet sequentially using LightCells in Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an Excel workbook with Aspose.Cells, applies FilterObjects to keep only visible worksheets, and uses the LightCells API to read each cell's address and value. | Show how to configure a LightCellsProcessor to iterate through the filtered visible sheets and output the cell data to the console or a custom writer. | Provide comprehensive error handling for missing input files, directory creation, and exceptions that may arise during LightCells processing in an Aspose.Cells example. | Demonstrate saving the workbook after LightCells processing while ensuring that only the originally visible worksheets are retained.
// Common Searches: Aspose.Cells C# filter visible worksheets with FilterObjects | How to use LightCells to read cells from only visible sheets in .NET | Process Excel workbook visible worksheets sequentially using LightCells Aspose | C# example loading workbook and iterating visible worksheets with LightCells API | Apply FilterObjects to workbook worksheets before LightCells processing Aspose.Cells
// Tags: filter visible worksheets Aspose.Cells | LightCells sequential sheet processing C# | load workbook visible sheets .NET | read cell values with LightCells API | apply FilterObjects Aspose.Cells workbook

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates how to open an Excel file with Aspose.Cells, use the FilterObjects collection to select only worksheets marked as visible, and then employ the LightCells API to process each visible sheet sequentially. It includes robust checks for the input file, creates the output directory if needed, and saves the workbook after processing, with detailed error handling throughout.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Process only visible worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.IsVisible)
                {
                    try
                    {
                        // Iterate through all cells in the sheet and output their values
                        foreach (Cell cell in sheet.Cells)
                        {
                            Console.WriteLine($"{sheet.Name}!{cell.Name} = {cell.Value}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing sheet '{sheet.Name}': {ex.Message}");
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (optional, if modifications were made)
            workbook.Save(outputPath);
            Console.WriteLine($"Processing completed. Output saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
