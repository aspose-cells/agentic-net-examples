// Title: Find the highest numeric value in an Excel worksheet and write it to a cell using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that opens an .xlsx file, scans every cell in the first worksheet, and returns the maximum numeric value. | Modify the program to insert the computed maximum into cell A1 and save the workbook as a new file. | Add robust error handling to report a missing input file and catch runtime exceptions in an Aspose.Cells script.
// Common Searches: c# aspose.cells get maximum numeric value from worksheet | how to write a calculated value back to a specific cell using Aspose.Cells | asp.net read all cells in Excel and find highest number with Aspose.Cells | saving modified Excel workbook after processing cells with Aspose.Cells
// Tags: Aspose.Cells maximum numeric value | C# iterate worksheet cells Aspose.Cells | write computed value to cell Aspose.Cells | load and save Excel workbook Aspose.Cells .NET | error handling missing file Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// // Loads "input.xlsx" with Aspose.Cells, iterates every cell in the first worksheet to locate the largest numeric entry, writes that value into cell A1, and saves the result as "output.xlsx" while handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            double maxNumeric = double.MinValue;

            // Iterate over all cells in the worksheet
            foreach (Cell cell in sheet.Cells)
            {
                // Consider only numeric cells
                if (cell.Type == CellValueType.IsNumeric)
                {
                    double value = cell.DoubleValue;
                    if (value > maxNumeric)
                    {
                        maxNumeric = value;
                    }
                }
            }

            // If no numeric cells were found, set max to 0 (or handle as needed)
            if (maxNumeric == double.MinValue)
            {
                maxNumeric = 0;
            }

            // Output the maximum numeric value
            Console.WriteLine("Maximum numeric value: " + maxNumeric);

            // Optionally write the result back into the worksheet
            sheet.Cells["A1"].PutValue(maxNumeric);

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
