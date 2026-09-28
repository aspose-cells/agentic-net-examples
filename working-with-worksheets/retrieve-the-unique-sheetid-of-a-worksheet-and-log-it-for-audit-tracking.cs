// Title: How to retrieve and log a worksheet's unique SheetId (index) using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an Excel workbook with Aspose.Cells, reads the Index property of a specified worksheet, and writes the worksheet name and index to the console with proper error handling. | Provide a snippet that checks for the existence of an .xlsx file, loads it via Aspose.Cells, obtains the worksheet's unique identifier (Index), and logs it for audit tracking.
// Common Searches: Aspose.Cells C# get worksheet index for audit logging | How to obtain sheet identifier in .NET Excel workbook using Aspose | Log worksheet unique ID with Aspose.Cells when file may be missing | Retrieve worksheet Index property in C# Aspose.Cells example
// Tags: Aspose.Cells get worksheet index | C# log worksheet identifier | audit worksheet IDs Aspose.Cells | handle missing Excel file Aspose.Cells | retrieve worksheet unique id .NET

using Aspose.Cells;
using System;
using System.IO;

// C# program that verifies an .xlsx file exists, loads it with Aspose.Cells, accesses a worksheet, reads its Index property (the unique SheetId within the workbook), and writes the worksheet name and index to the console while handling exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (or any specific one you need)
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the worksheet's index, which serves as its unique identifier within the workbook
            int sheetIndex = worksheet.Index;

            // Log the index for audit tracking
            Console.WriteLine($"Worksheet '{worksheet.Name}' has Index: {sheetIndex}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
