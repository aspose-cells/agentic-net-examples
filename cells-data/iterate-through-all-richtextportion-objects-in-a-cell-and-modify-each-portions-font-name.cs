// Title: Set the font of a specific cell to Arial using Aspose.Cells for .NET
// Common Searches: how to change the font of a single cell to Arial with Aspose.Cells C# | Aspose.Cells set cell style font name programmatically | C# example for applying Arial font to cell A1 in Excel workbook | using Aspose.Cells to modify cell font and save workbook
// Tags: set cell font name Aspose.Cells | apply style to specific cell .NET | change cell font to Arial C# | Aspose.Cells workbook save after style change

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel workbook, verifies the input file, accesses the first worksheet and cell A1, changes the cell's font to Arial by updating its style, saves the modified workbook to a new file, and includes basic exception handling.
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
                Console.WriteLine($"Input file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Target cell (e.g., A1)
            Cell cell = worksheet.Cells["A1"];

            // Change the font of the entire cell to Arial
            Style style = cell.GetStyle();
            style.Font.Name = "Arial";
            cell.SetStyle(style);

            // Save the workbook with the changes
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
