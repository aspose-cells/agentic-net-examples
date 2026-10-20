// Title: How to delete a named range and refresh dependent formulas in an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# delete a defined name and recalculate workbook formulas | remove named range from Excel using Aspose.Cells and update dependent cells | call workbook.CalculateFormula after removing a named range in .NET | C# code to delete Excel named range and refresh formulas with Aspose.Cells | Aspose.Cells example for deleting a name and recalculating formulas
// Tags: Aspose.Cells delete named range C# | Workbook.CalculateFormula usage .NET | update dependent formulas after name removal | Excel named range removal Aspose.Cells | recalculate formulas after name deletion

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file, checks for a named range called "MyRange", removes it from the workbook's NameCollection if present, invokes workbook.CalculateFormula to update all formulas that referenced the deleted range, and saves the modified workbook. It includes basic file existence checks and exception handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string rangeNameToDelete = "MyRange";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Access the collection of defined names (named ranges) in the workbook
            NameCollection namedRanges = workbook.Worksheets.Names;

            // Check if the named range exists and remove it
            if (namedRanges[rangeNameToDelete] != null)
            {
                namedRanges.Remove(rangeNameToDelete);
                Console.WriteLine($"Named range '{rangeNameToDelete}' removed.");
            }
            else
            {
                Console.WriteLine($"Named range '{rangeNameToDelete}' not found.");
            }

            // Recalculate all formulas to update dependent cells
            workbook.CalculateFormula();

            // Save the updated workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
