// Title: Replace #NAME? formula errors with #REF! in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that opens a workbook, recalculates all formulas, locates cells whose evaluated value is "#NAME?", and changes those values to "#REF!" before saving the file. | Create a reusable method `FixUndefinedNameErrors(Workbook workbook)` that scans every worksheet, replaces any "#NAME?" results with "#REF!", and returns a list of the affected cell addresses. | Update the sample program to log each cell address where a "#NAME?" error is converted to "#REF!" and display the log on the console prior to saving the workbook.
// Common Searches: Aspose.Cells C# replace #NAME? error with #REF! in Excel file | detect undefined name formula errors using Aspose.Cells .NET | iterate over used range cells to change #NAME? to #REF! in a workbook | how to programmatically convert #NAME? to #REF! with Aspose.Cells
// Tags: replace #NAME? with #REF! Aspose.Cells | detect undefined name formula errors .NET | iterate worksheet cells Aspose.Cells | recalculate formulas before error handling C# | save modified Excel workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, forces a full formula calculation, walks through each worksheet's used range, identifies cells whose formula result is the "#NAME?" error, replaces that result with the "#REF!" error string, and saves the updated workbook, while handling file I/O and cell‑level exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            Workbook workbook;
            try
            {
                // Load the workbook
                workbook = new Workbook(inputPath);
            }
            catch (Exception loadEx)
            {
                Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
                return;
            }

            // Calculate all formulas so that error values are up‑to‑date
            workbook.CalculateFormula();

            // Process each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Limit iteration to the used range of the sheet
                Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;

                foreach (Cell cell in usedRange)
                {
                    try
                    {
                        // Only examine cells that contain a formula
                        if (!cell.IsFormula)
                            continue;

                        // After calculation, error values appear as strings like "#NAME?"
                        if (cell.Value is string val && val.Equals("#NAME?", StringComparison.OrdinalIgnoreCase))
                        {
                            // Replace #NAME? with #REF! error representation
                            cell.PutValue("#REF!");
                        }
                    }
                    catch (Exception exCell)
                    {
                        // Log cell‑level issues but continue processing
                        Console.WriteLine($"Error processing cell {cell.Name}: {exCell.Message}");
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
