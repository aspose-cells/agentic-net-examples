// Title: How to delete every shape from an Excel worksheet with Aspose.Cells for .NET and then export it as CSV
// AI Prompts: Write C# code using Aspose.Cells that removes all shapes from a specified worksheet and saves the result directly to a CSV file. | Show the pattern for iterating backwards through Worksheet.Shapes to safely delete each shape before calling Workbook.Save with SaveFormat.Csv. | Create a complete example that verifies the source .xlsx file exists, clears all drawings from the first sheet, and outputs a clean CSV without any graphics.
// Common Searches: Aspose.Cells .NET delete all shapes from a worksheet before CSV export | How to remove pictures and charts from Excel file using Aspose.Cells and save as CSV | C# code to clear drawings in a worksheet then convert to CSV with Aspose.Cells | Export Excel to CSV without embedded objects using Aspose.Cells for .NET
// Tags: remove worksheet shapes Aspose.Cells C# | clear drawings before CSV export Aspose.Cells | reverse iteration shape removal Aspose.Cells | save worksheet as CSV after shape deletion Aspose.Cells | Aspose.Cells CSV export without graphics

using Aspose.Cells;
using System;
using System.IO;

// Loads an Excel workbook, iterates backwards to delete every shape from the first worksheet, and saves the cleaned worksheet as a CSV file, including a file‑existence check and exception handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.csv";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (or specify by name)
            Worksheet worksheet = workbook.Worksheets[0];

            // Remove all shapes from the worksheet (iterate backwards)
            for (int i = worksheet.Shapes.Count - 1; i >= 0; i--)
            {
                worksheet.Shapes.RemoveAt(i);
            }

            // Save the cleaned worksheet as CSV
            workbook.Save(outputPath, SaveFormat.Csv);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
