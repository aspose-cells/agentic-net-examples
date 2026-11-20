// Title: Load an Excel workbook from a MemoryStream, modify a cell formula, and recalculate all formulas with custom CalculationOptions using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# load workbook from MemoryStream and recalculate formulas with options | How to change a cell formula and run CalculateFormula with custom CalculationOptions in .NET | Update Excel formula programmatically and ignore errors during calculation using Aspose.Cells | Recalculate dependent formulas recursively after modifying a workbook in C# | Save modified workbook to a new file after applying CalculationOptions with Aspose.Cells
// Tags: memory stream workbook loading Aspose.Cells | modify cell formula C# Aspose.Cells | custom calculationoptions for formula evaluation | recursive formula recalculation Aspose.Cells | ignore errors during Excel calculation Aspose.Cells | save workbook to file after calculation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example reads an existing XLSX file into a byte array, loads it into an Aspose.Cells Workbook via a MemoryStream, changes the formula in cell A1 to "=SUM(B1:B10)", sets CalculationOptions to ignore errors and recalculate dependent formulas recursively, invokes Workbook.CalculateFormula with those options, and finally saves the updated workbook to a new file using another MemoryStream, with basic file‑existence checks and exception handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Read the workbook bytes.
            byte[] workbookData = File.ReadAllBytes(inputPath);

            // Load the workbook from a memory stream.
            using (MemoryStream inputStream = new MemoryStream(workbookData))
            {
                Workbook workbook = new Workbook(inputStream);

                // Change the formula of cell A1 in the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];
                Cell targetCell = sheet.Cells["A1"];
                targetCell.Formula = "=SUM(B1:B10)";

                // Set calculation options.
                CalculationOptions calcOptions = new CalculationOptions
                {
                    // Example options – adjust as needed.
                    IgnoreError = true,   // Ignore errors during calculation.
                    Recursive = true      // Recalculate dependent formulas recursively.
                };

                // Recalculate all formulas using the specified options.
                workbook.CalculateFormula(calcOptions);

                // Save the modified workbook to a new memory stream and then to a file.
                using (MemoryStream outputStream = new MemoryStream())
                {
                    workbook.Save(outputStream, SaveFormat.Xlsx);
                    File.WriteAllBytes(outputPath, outputStream.ToArray());
                }

                Console.WriteLine($"Workbook processed successfully. Output saved to \"{outputPath}\".");
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
