// Title: How to loop through a worksheet's used range and log each cell's address and value with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an Excel workbook with Aspose.Cells, obtains the first worksheet's MaxDisplayRange, and writes each cell's address and its value to the console. | Create a reusable C# method that takes an Excel file path, iterates over all cells in the first sheet's used range using Aspose.Cells, and returns a collection of strings formatted as "A1: 123".
// Common Searches: Aspose.Cells C# iterate over MaxDisplayRange and print cell names | C# read every cell value from first worksheet using Aspose.Cells | How to get cell address and value while looping through used range in Aspose.Cells | Example of logging Excel cell addresses with Aspose.Cells .NET
// Tags: Aspose.Cells used range iteration C# | Aspose.Cells cell address logging | MaxDisplayRange cell enumeration Aspose | C# console output Excel cell values | read worksheet cells Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The program loads "input.xlsx", accesses the first worksheet's MaxDisplayRange via Aspose.Cells, iterates over each cell, prints the cell name and its value to the console, and finally saves the workbook as "output.xlsx".
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get the used range of the worksheet (explicitly use Aspose.Cells.Range)
            Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;

            // Iterate through each cell in the used range
            for (int row = usedRange.FirstRow; row < usedRange.FirstRow + usedRange.RowCount; row++)
            {
                for (int col = usedRange.FirstColumn; col < usedRange.FirstColumn + usedRange.ColumnCount; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    // Log cell address and its value
                    Console.WriteLine($"{cell.Name}: {cell.Value}");
                }
            }

            // Save the workbook (optional)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
