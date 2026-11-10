// Title: Add a worksheet to an existing Excel file and save it as a new .xlsx while preserving custom properties with Aspose.Cells for .NET (C#)
// Common Searches: C# how to add a new sheet to an existing workbook and keep custom properties using Aspose.Cells | Aspose.Cells .NET save modified workbook to a new .xlsx without losing XML parts | preserve custom XML parts when saving Excel file after adding worksheet in C# | save workbook to a different file name while retaining metadata with Aspose.Cells | load Excel file, add sheet, and export to new file preserving properties Aspose.Cells C#
// Tags: add worksheet Aspose.Cells C# | save workbook as new xlsx Aspose.Cells | keep workbook metadata Aspose.Cells | save with XML parts intact Aspose.Cells | modify existing Excel file Aspose.Cells example

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an existing Excel workbook (input.xlsx) using Aspose.Cells for .NET, adds a new worksheet named "NewSheet", writes a value to cell A1, and then saves the updated workbook to a new file (output.xlsx) in Xlsx format, ensuring that all custom properties and XML parts are retained.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Add a new worksheet and write a value
            int newSheetIndex = workbook.Worksheets.Add();
            Worksheet newSheet = workbook.Worksheets[newSheetIndex];
            newSheet.Name = "NewSheet";
            newSheet.Cells["A1"].PutValue("Hello Aspose!");

            // Save the modified workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
