// Title: How to copy only the cell formatting from A1:B5 to C1:D5 in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells to copy the style of cells A1:B5 to the range C1:D5 without transferring the cell values. | Implement a C# routine that loads an Excel workbook, creates source and destination ranges, and applies Range.CopyStyle to move only formatting. | Write code that opens input.xlsx, copies formatting from rows 1‑5, columns A‑B to columns C‑D of the same rows, and saves the result as output.xlsx.
// Common Searches: Aspose.Cells C# copy only cell style from one range to another | Range.CopyStyle example for copying formatting in a .NET Excel workbook | How to transfer formatting without data between A1:B5 and C1:D5 using Aspose.Cells
// Tags: Aspose.Cells Range.CopyStyle usage | copy cell formatting Excel .NET | transfer styles between ranges C# | load workbook and apply formatting Aspose.Cells | Excel range formatting copy without values

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;
using System.Drawing;

// Loads input.xlsx (creates a sample file if missing), defines source range A1:B5 and destination range C1:D5, copies only the formatting using Range.CopyStyle, and saves the modified workbook to output.xlsx.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists; create a simple workbook if it does not.
            if (!File.Exists(inputPath))
            {
                var tempWb = new Workbook();
                var tempWs = tempWb.Worksheets[0];

                // Populate sample data in A1:B5 and apply a basic style.
                var style = tempWb.CreateStyle();
                style.ForegroundColor = Color.LightBlue;
                style.Pattern = BackgroundType.Solid;

                for (int row = 0; row < 5; row++)
                {
                    tempWs.Cells[row, 0].PutValue($"A{row + 1}");
                    tempWs.Cells[row, 1].PutValue($"B{row + 1}");
                    tempWs.Cells[row, 0].SetStyle(style);
                    tempWs.Cells[row, 1].SetStyle(style);
                }

                tempWb.Save(inputPath);
            }

            // Load the workbook from the input file.
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet.
            Worksheet worksheet = workbook.Worksheets[0];

            // Define source (A1:B5) and destination (C1:D5) ranges.
            AsposeRange sourceRange = worksheet.Cells.CreateRange("A1", "B5");
            AsposeRange destinationRange = worksheet.Cells.CreateRange("C1", "D5");

            // Copy only the formatting (styles) from source to destination.
            sourceRange.CopyStyle(destinationRange);

            // Save the modified workbook.
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook processed successfully. Output saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
