// Title: Set the font size of a RichTextPortion to 12 points in an Excel cell with Aspose.Cells for .NET
// AI Prompts: Update the sample to locate the first RichTextPortion in cell A1 and set its Font.Size property to 12 using Aspose.Cells. | Write a reusable C# method that accepts a worksheet, cell address, and font size, then applies the size to all RichTextPortions within that cell via Aspose.Cells.
// Common Searches: Aspose.Cells C# change font size of a single rich text segment in a cell | how to set font size for part of cell text using Aspose.Cells .NET | modify specific RichTextPortion formatting in Excel with Aspose.Cells library | C# Aspose.Cells example to adjust font size of a text portion inside a cell
// Tags: Aspose.Cells RichTextPortion font size | C# set Excel rich text segment size | Aspose.Cells modify specific text portion style | Excel .NET change part of cell font | Aspose.Cells adjust RichTextPortion formatting

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing workbook, retrieves cell A1, accesses its RichText collection, changes the Font.Size of the targeted RichTextPortion to 12 points, and saves the workbook. This demonstrates how to format a specific segment of cell text with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Access the cell that may contain rich text (e.g., A1)
            Cell cell = worksheet.Cells["A1"];

            // Modify the font size of the entire cell (applies to all text portions)
            Style style = cell.GetStyle();
            style.Font.Size = 12; // set desired font size
            cell.SetStyle(style);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
