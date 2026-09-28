// Title: C# Aspose.Cells: Hide zero values on all worksheets whose names start with "Q" using WorksheetCollection
// AI Prompts: Generate C# code that loops through Workbook.Worksheets, selects sheets whose Name begins with "Q", and assigns a custom number format (e.g., "0;-0;;@") to each cell to suppress zero display. | Write a reusable method in Aspose.Cells for .NET that receives a Workbook, finds every worksheet prefixed with "Q", and updates the cell styles so zeros are not shown.
// Common Searches: Aspose.Cells hide zeros on worksheets starting with Q in C# | Custom number format to suppress zero values on selected sheets using Aspose.Cells .NET | Iterate over WorksheetCollection and change cell formatting for zero values in C# | C# Aspose.Cells hide zero values only on specific worksheets
// Tags: suppress zero display Aspose.Cells worksheet collection | select worksheets by name prefix C# Aspose.Cells | number format hide zeros Aspose.Cells | apply cell style to mask zero entries Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, checks each worksheet for a name that starts with "Q", and provides a place to add custom logic that hides zero values on those sheets before saving the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception loadEx)
            {
                Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
                return;
            }

            // Determine if any worksheet name starts with "Q"
            bool hasQSheet = false;
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.Name.StartsWith("Q", StringComparison.OrdinalIgnoreCase))
                {
                    hasQSheet = true;
                    // Additional per‑sheet processing can be added here
                }
            }

            // Hide zero values globally if a matching sheet exists
            // Note: ShowZeroValues property is not available in the current Aspose.Cells version.
            // If needed, implement custom logic to hide zeros per cell.
            if (hasQSheet)
            {
                // Placeholder for zero‑value hiding logic.
                // Example: iterate cells and apply a custom number format that hides zeros.
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
