// Title: Determine if an Excel worksheet contains merged cells using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, accesses a worksheet, and uses GetMergedAreas to find out if any merged cells are present. | Show how to safely check for merged cell ranges in a worksheet, including file existence verification and exception handling, with Aspose.Cells for .NET. | Provide a concise example that returns a boolean indicating the presence of merged cells in the first worksheet of an Excel file using Aspose.Cells.
// Common Searches: C# Aspose.Cells example to check for merged cells in a worksheet | How to use GetMergedAreas to detect merged ranges in an Excel file with Aspose.Cells | Determine if an Excel sheet has any merged cells before processing using Aspose.Cells .NET | Aspose.Cells GetMergedAreas returns empty array when no merged cells are present | Validate merged cells existence in workbook with Aspose.Cells C# code
// Tags: merged cells detection Aspose.Cells | GetMergedAreas usage C# | worksheet merged range check .NET | validate merged cells before processing | file existence handling Aspose.Cells | exception handling workbook load

using System;
using System.IO;
using Aspose.Cells;

// The program loads an Excel file, verifies its existence, accesses the first worksheet, retrieves merged cell areas via GetMergedAreas, and prints whether any merged cells are found, with proper exception handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve all merged cell ranges in the worksheet using the recommended API
            CellArea[] mergedRanges = worksheet.Cells.GetMergedAreas();

            // Determine if any merged cells exist
            bool hasMergedCells = mergedRanges != null && mergedRanges.Length > 0;

            // Output the result
            Console.WriteLine(hasMergedCells
                ? "The worksheet contains merged cells."
                : "The worksheet does not contain any merged cells.");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
