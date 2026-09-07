// Title: Create a new worksheet, write a value to cell A1, and save the workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write a C# console program that opens an existing .xlsx file with Aspose.Cells, adds a worksheet named "DemoSheet", inserts the text "Hello Aspose.Cells!" into cell A1, and saves the modified workbook as a new file. | Extend the program to read data from a CSV file and populate the newly added worksheet row by row using Aspose.Cells, then save the result.
// Common Searches: how to add a new worksheet to an existing Excel file using Aspose.Cells in C# | asp.net c# write text to cell A1 with Aspose.Cells and export as new workbook | load workbook with Aspose.Cells and check file existence before saving | c# catch FileNotFoundException when opening an Excel file with Aspose.Cells
// Tags: add worksheet Aspose.Cells .NET | write cell value Aspose.Cells C# | save modified workbook Aspose.Cells | load existing Excel file Aspose.Cells | handle file not found Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example verifies that input.xlsx exists, loads it with Aspose.Cells, adds a worksheet called "DemoSheet", writes "Hello Aspose.Cells!" into cell A1, and saves the updated workbook as output.xlsx while handling any runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the input Excel file.
            string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the input file.
            Workbook workbook = new Workbook(inputPath);

            // Add a new worksheet and write a sample value.
            int newSheetIndex = workbook.Worksheets.Add();
            Worksheet newSheet = workbook.Worksheets[newSheetIndex];
            newSheet.Name = "DemoSheet";
            newSheet.Cells["A1"].PutValue("Hello Aspose.Cells!");

            // Define the output file path.
            string outputPath = "output.xlsx";

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook successfully saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display the error message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
