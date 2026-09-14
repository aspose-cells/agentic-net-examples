// Title: How to retrieve a worksheet by its name using Aspose.Cells for .NET (C#)
// AI Prompts: Write a C# method that accepts an Excel file path and a sheet name, loads the workbook with Aspose.Cells, returns the Worksheet object, and throws a descriptive exception if the sheet does not exist. | Generate C# code that opens a workbook, accesses the worksheet named "MySheetName", verifies that the worksheet is present, and stores the Worksheet reference for further manipulation.
// Common Searches: Aspose.Cells C# get worksheet object by sheet name | How to verify worksheet existence before accessing it with Aspose.Cells .NET | Retrieve specific sheet from an Excel workbook using Aspose.Cells in C# | C# Aspose.Cells example for handling missing worksheet | Access worksheet by name and store reference for further processing Aspose.Cells
// Tags: retrieve worksheet by name Aspose.Cells | load workbook and access specific sheet C# | check worksheet existence Aspose.Cells .NET | handle missing worksheet exception C# | worksheet reference for further processing Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads 'input.xlsx' with Aspose.Cells, attempts to fetch the worksheet named 'MySheetName' via the Worksheets collection, checks for a null result to handle a missing sheet, prints the accessed sheet name, and includes basic error handling.
class WorksheetAccessExample
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Attempt to get the worksheet by name
            string targetSheetName = "MySheetName";
            Worksheet sheet = workbook.Worksheets[targetSheetName];

            if (sheet == null)
            {
                Console.WriteLine($"Worksheet \"{targetSheetName}\" not found in the workbook.");
                return;
            }

            // Example operation: read the name to verify access
            string sheetName = sheet.Name;
            Console.WriteLine($"Accessed worksheet: {sheetName}");

            // Optional: Save the workbook after any modifications
            // string outputPath = "output.xlsx";
            // workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
