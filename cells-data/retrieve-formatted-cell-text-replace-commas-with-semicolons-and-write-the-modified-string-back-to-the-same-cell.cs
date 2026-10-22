// Title: Read formatted text from a specific Excel cell, replace commas with semicolons, and write the result back using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to read the displayed value of cell A1, swap commas for semicolons, and save the workbook. | Demonstrate how to loop through a column in an Excel sheet and perform a comma‑to‑semicolon conversion on each cell's formatted text using Aspose.Cells in .NET. | Create a C# example with error handling that checks for the source file, modifies a cell's formatted string, and writes the updated workbook to a new file using Aspose.Cells.
// Common Searches: Aspose.Cells read formatted cell text and replace commas in C# | C# Aspose.Cells change commas to semicolons in Excel cell | How to modify cell display value with Aspose.Cells .NET | Replace character in Excel cell using Aspose.Cells library | Save workbook after editing cell text with Aspose.Cells C#
// Tags: Aspose.Cells get formatted cell value | Aspose.Cells comma to semicolon conversion | Aspose.Cells write updated text to cell | Aspose.Cells save workbook after edit | C# Excel cell text manipulation Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads 'input.xlsx', reads the displayed value of cell A1, substitutes commas with semicolons, writes the modified string back to the same cell, and saves the workbook as 'output.xlsx' using Aspose.Cells for .NET.
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Target cell (e.g., A1)
            Cell cell = worksheet.Cells["A1"];

            // Get the cell's displayed (formatted) text
            string formattedText = cell.StringValue; // equivalent to GetStringValue with default strategy

            // Replace commas with semicolons
            string modifiedText = formattedText.Replace(",", ";");

            // Write the modified string back to the same cell
            cell.PutValue(modifiedText);

            // Save the workbook to the output file (overwrites if exists)
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
