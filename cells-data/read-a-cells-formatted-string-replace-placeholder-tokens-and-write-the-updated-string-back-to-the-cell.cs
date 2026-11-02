// Title: Read a cell's displayed string, replace {Name} and {Date} placeholders, and write the updated value back with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells get displayed text of a cell and modify it | replace placeholder tokens in Excel cell using Aspose.Cells GetStringValue | write updated string back to same cell Aspose.Cells PutValue example | how to use CellValueFormatStrategy.DisplayString in Aspose.Cells C# | save modified workbook after editing cell with Aspose.Cells
// Tags: Cell.GetStringValue DisplayString usage | Aspose.Cells placeholder token replacement | Cell.PutValue update string example | Workbook.Save after cell edit | read formatted cell text Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook, reads the displayed text of cell B2 using GetStringValue with the DisplayString strategy, replaces the {Name} and {Date} placeholders, writes the new string back to the same cell with PutValue, and saves the workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            // Define input and output file paths
            string inputPath = "Input.xlsx";
            string outputPath = "Output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (index 0)
            Worksheet sheet = workbook.Worksheets[0];

            // Specify the cell address (e.g., "B2")
            Cell cell = sheet.Cells["B2"];

            // Get the cell's formatted (display) string using the appropriate overload
            string formattedText = cell.GetStringValue(CellValueFormatStrategy.DisplayString);

            // Replace placeholder tokens with actual values
            string updatedText = formattedText
                .Replace("{Name}", "John")
                .Replace("{Date}", DateTime.Today.ToShortDateString());

            // Write the updated string back to the same cell
            cell.PutValue(updatedText);

            // Save the modified workbook to the output file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
