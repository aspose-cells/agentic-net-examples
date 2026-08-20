// Title: C# Example: Remove All Table Slicers from a Worksheet Using Aspose.Cells
// Description: Loads a workbook, checks for tables on the first worksheet, iterates the SlicerCollection in reverse, removes every slicer, and saves the cleaned file. Ideal for simplifying worksheets before distribution or reuse.
// Keywords: Aspose.Cells C# remove slicers | delete worksheet slicers .NET | clear slicer collection Aspose | remove table slicers programmatically | Aspose.Cells example GitHub | C# workbook cleanup | Excel slicer removal code
// Common Searches: how to delete all slicers with Aspose.Cells C# | remove slicers from a specific table using Aspose.Cells | Aspose.Cells C# clear slicer collection example | C# code to remove worksheet slicers Aspose | Aspose.Cells remove slicers before saving workbook
// Developer Intent: Eliminate every slicer on the target worksheet to provide a clean, uncluttered interface.
// Use Cases: Prepare a template workbook for reuse by stripping out all slicers added during prior analysis. | Clean up a report before sharing with stakeholders, ensuring no interactive slicers remain. | Reset the UI after dynamically modifying tables, preventing orphaned slicers from confusing users.
// AI Prompts: Generate C# code with Aspose.Cells that removes only slicers linked to a given table name while preserving others. | Show how to log the names of removed slicers and keep slicers on other worksheets untouched. | Provide a reusable method that accepts a worksheet and optional table identifier to delete matching slicers.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;

// Loads a workbook, checks for tables on the first worksheet, iterates the SlicerCollection in reverse, removes every slicer, and saves the cleaned file. Ideal for simplifying worksheets before distribution or reuse.
class RemoveTableSlicers
{
    static void Main()
    {
        const string inputPath = "InputWorkbook.xlsx";
        const string outputPath = "OutputWorkbook.xlsx";

        try
        {
            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one table
            if (worksheet.ListObjects.Count == 0)
            {
                Console.WriteLine("No tables found in the worksheet.");
                return;
            }

            // Access the slicer collection on the worksheet
            SlicerCollection slicers = worksheet.Slicers;

            // Iterate backwards to safely remove slicers
            for (int i = slicers.Count - 1; i >= 0; i--)
            {
                // Remove each slicer (or add custom logic to filter specific slicers)
                slicers.RemoveAt(i);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
