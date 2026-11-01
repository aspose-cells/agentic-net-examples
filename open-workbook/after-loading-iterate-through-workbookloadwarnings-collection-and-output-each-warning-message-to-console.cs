// Title: How to iterate Workbook.LoadWarnings and print each warning to the console using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells read workbook load warnings and display them | Iterate over LoadWarnings collection after opening an Excel file with Aspose.Cells | Print Aspose.Cells workbook load warning messages to console in .NET | How to handle load warnings when loading an .xlsx file with Aspose.Cells in C# | Retrieve and log workbook load warnings using Aspose.Cells API
// Tags: Aspose.Cells iterate LoadWarnings | C# workbook load warnings console output | Aspose.Cells warning messages extraction | handling load warnings .NET Excel | Aspose.Cells workbook loading diagnostics

using System;
using System.IO;
using Aspose.Cells;

// The example verifies that the input.xlsx file exists, creates an empty workbook if it does not, and then loads the workbook with Aspose.Cells. After loading, it iterates through the Workbook.LoadWarnings collection and writes each warning message to the console, all wrapped in a try‑catch block to capture unexpected errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        try
        {
            // Ensure the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"File not found: {inputPath}");
                // Create a new empty workbook to continue execution
                var emptyWorkbook = new Workbook();
                emptyWorkbook.Save(inputPath);
                Console.WriteLine($"Created empty workbook at: {inputPath}");
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // (Optional) Add any further processing of the workbook here
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display the error message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
