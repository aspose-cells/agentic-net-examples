// Title: Convert all formula cells to their calculated values and save an Excel workbook as CSV using Aspose.Cells for .NET (C#)
// AI Prompts: Load an .xlsx workbook with Aspose.Cells, call workbook.CalculateFormula(), and then export the result to a .csv file in C#. | Using Aspose.Cells in C#, evaluate every formula in a spreadsheet and write the evaluated values to a CSV output.
// Common Searches: asp.net aspose.cells calculate all formulas before CSV export | c# save excel as csv with evaluated formula results using Aspose.Cells | how to force formula calculation when converting xlsx to csv in .NET
// Tags: calculate formulas Aspose.Cells | save workbook as CSV Aspose.Cells | evaluate formulas before CSV export C# | convert formula cells to values Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing XLSX file, forces calculation of all formulas with workbook.CalculateFormula(), and then saves the workbook as a CSV file using Aspose.Cells for .NET, including file existence checking and error handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.csv";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Calculate all formulas so that the latest values are stored
            workbook.CalculateFormula();

            // Save the workbook as CSV. By default, formulas are saved as their calculated values.
            workbook.Save(outputPath, SaveFormat.Csv);

            Console.WriteLine($"Workbook successfully saved as CSV to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
