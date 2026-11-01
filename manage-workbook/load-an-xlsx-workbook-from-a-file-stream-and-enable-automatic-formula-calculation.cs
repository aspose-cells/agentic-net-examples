// Title: Load an XLSX workbook from a FileStream and calculate all formulas with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells load workbook from FileStream and recalculate formulas | How to open an Excel file via stream and trigger CalculateFormula in Aspose.Cells | Aspose.Cells .NET example for reading XLSX from FileStream and saving after formula evaluation | Using Aspose.Cells to read Excel file from memory stream and automatically compute formulas | Error handling when loading an Excel workbook from a stream with Aspose.Cells in C#
// Tags: filestream workbook loading Aspose.Cells | automatic formula calculation Aspose.Cells | saving workbook after calculation Aspose.Cells | input file existence verification C# | exception handling for Aspose.Cells file operations

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies that the input XLSX file exists, opens it through a FileStream, loads it into an Aspose.Cells Workbook, runs full formula calculation with CalculateFormula(), and then saves the updated workbook to a new file while handling potential exceptions.
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
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook from a file stream
            using (FileStream fileStream = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
            {
                Workbook workbook = new Workbook(fileStream);

                // Calculate all formulas in the workbook
                workbook.CalculateFormula();

                // Save the workbook to the specified output file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
