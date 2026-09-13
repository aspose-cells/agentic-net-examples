// Title: How to delete a named range and recalculate dependent formulas in an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Generate C# code that checks if a named range exists in a workbook, removes it, calls Workbook.CalculateFormula, and saves the file using Aspose.Cells. | Show a step‑by‑step example of loading an .xlsx file, deleting a specific named range, triggering full formula recalculation, and writing the result with Aspose.Cells in a .NET console application. | Provide a minimal Aspose.Cells snippet that removes a range named "MyRange" and ensures all formulas referencing that range are updated by invoking CalculateFormula.
// Common Searches: aspnet remove named range from Excel file and recalculate formulas | c# Aspose.Cells delete named range then update dependent cells | how to force formula recalculation after deleting a named range in Aspose.Cells | sample code for Workbook.CalculateFormula after range removal | checking existence of a named range before deletion using Aspose.Cells
// Tags: delete named range Aspose.Cells C# | Workbook.CalculateFormula after range removal | named range existence check Aspose.Cells | recalculate dependent formulas .NET Excel library | Aspose.Cells remove range and update formulas

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example loads an existing .xlsx workbook, verifies whether a named range called "MyRange" exists, removes that range from the workbook's Names collection, invokes Workbook.CalculateFormula to refresh any formulas that depended on the deleted range, and saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string rangeName = "MyRange";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the input file
            var workbook = new Workbook(inputPath);

            // Access the collection of named ranges
            var names = workbook.Worksheets.Names;
            var name = names[rangeName];

            if (name != null)
            {
                // Retrieve the actual Range object using the name
                AsposeRange namedRange = workbook.Worksheets.GetRangeByName(rangeName);
                // (namedRange can be used here if further processing is needed)

                // Remove the named range from the workbook
                names.Remove(rangeName);
                Console.WriteLine($"Named range \"{rangeName}\" removed.");
            }
            else
            {
                Console.WriteLine($"Named range \"{rangeName}\" does not exist.");
            }

            // Recalculate all formulas so that dependent cells are updated
            workbook.CalculateFormula();

            // Save the modified workbook to the output file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
