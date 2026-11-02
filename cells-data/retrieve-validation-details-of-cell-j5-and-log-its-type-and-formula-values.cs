// Title: How to read and log a cell’s data‑validation type and formulas (J5) with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# read data validation type of cell J5 | Get Validation.Formula1 and Validation.Formula2 from a specific cell using Aspose.Cells | Check if a worksheet cell has data validation and output its details in a .NET console app | Retrieve data‑validation settings for a cell in an existing Excel workbook with Aspose.Cells | Log Excel cell validation properties (type, formula1, formula2) using Aspose.Cells in C#
// Tags: Aspose.Cells read cell validation | C# retrieve validation type Aspose | Aspose.Cells extract validation formulas | Excel data validation extraction .NET | log validation details console Aspose

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, accesses the first worksheet, obtains the Validation object for cell J5, and writes the validation type along with Formula1 and Formula2 to the console. It demonstrates how to read and log data‑validation settings using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve cell J5
            Cell cell = worksheet.Cells["J5"];

            // Get the validation object associated with the cell
            Validation validation = cell.GetValidation();

            // Log validation details
            if (validation != null)
            {
                // Use the correct property name for validation type
                Console.WriteLine($"Validation Type: {validation.Type}");
                Console.WriteLine($"Formula1: {validation.Formula1}");
                Console.WriteLine($"Formula2: {validation.Formula2}");
            }
            else
            {
                Console.WriteLine("No validation set on cell J5.");
            }

            // Save the workbook (optional, only if changes were made)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
