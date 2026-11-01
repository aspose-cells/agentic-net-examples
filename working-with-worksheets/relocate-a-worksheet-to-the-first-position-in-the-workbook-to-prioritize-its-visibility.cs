// Title: How to move a worksheet to the first tab in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# move worksheet to first position in workbook | C# code to set worksheet index to 0 with Aspose.Cells | Reorder Excel sheets programmatically using Aspose.Cells .NET | Place specific worksheet at the beginning of an Excel file using Aspose.Cells | Move DataSheet worksheet to first tab Aspose.Cells example
// Tags: Worksheet.MoveTo usage Aspose.Cells | change worksheet order .NET Excel | worksheet position zero Aspose.Cells | Excel sheet reordering C# Aspose | prioritize sheet visibility Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel workbook, locates the worksheet named "DataSheet", moves it to the first tab (index 0) using the Worksheet.MoveTo method, and saves the updated file, while handling missing files and runtime exceptions.
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

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Identify the worksheet to relocate (by name)
            Worksheet sheetToMove = workbook.Worksheets["DataSheet"];
            if (sheetToMove == null)
            {
                Console.WriteLine("Worksheet 'DataSheet' not found.");
                return;
            }

            // Move the worksheet to the first position (index 0)
            sheetToMove.MoveTo(0);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
