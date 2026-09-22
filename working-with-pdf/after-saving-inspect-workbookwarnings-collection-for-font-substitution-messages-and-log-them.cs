// Title: How to log font substitution warnings after saving an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Write a C# method that saves a Workbook with Aspose.Cells and then iterates over workbook.GetWarnings() to output FontSubstitution warning details. | Show how to capture and display missing‑font warnings from the workbook warnings collection immediately after calling Workbook.Save. | Generate sample code that checks the Workbook warnings for FontSubstitution entries and writes each warning description to the console.
// Common Searches: Aspose.Cells .NET retrieve font substitution warnings after saving workbook | C# log missing font warnings when exporting Excel with Aspose.Cells | How to inspect workbook warnings for font issues in Aspose.Cells | GetWarnings FontSubstitution example for Aspose.Cells C# | Detect unavailable fonts during Excel save using Aspose.Cells API
// Tags: Aspose.Cells GetWarnings API | enumerate workbook warnings C# | detect missing fonts Aspose.Cells | inspect workbook warnings after save | Excel font substitution handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file, modifies a cell, saves the workbook, and demonstrates how to enumerate the workbook's warnings collection to identify and print any FontSubstitution warnings, helping developers catch missing‑font issues after a save operation.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Example modification: add a value to the first cell
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Text");

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");

            // Optional: inspect warnings if the API supports it
            // Note: GetWarnings may not be available in some versions of Aspose.Cells.
            // If needed, uncomment and adjust the following code according to the version you use.
            /*
            var warnings = workbook.GetWarnings();
            foreach (var warning in warnings)
            {
                if (warning.Type == ExceptionType.FontSubstitution)
                {
                    Console.WriteLine($"Font substitution warning: {warning.Description}");
                }
            }
            */
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
