// Title: Handle missing named range in Aspose.Cells for .NET with try‑catch and console logging
// AI Prompts: Write C# code that loads an Excel workbook using Aspose.Cells, attempts to retrieve a specific named range, verifies the result is not null, and logs a descriptive error message when the range is absent, wrapping the operation in a try‑catch block for unexpected exceptions. | Demonstrate how to surround workbook.Save with exception handling while also capturing and logging any errors that occur while accessing a named range in an Aspose.Cells workbook.
// Common Searches: Aspose.Cells C# how to check if a named range exists before using it | C# try-catch example for missing named range in Aspose.Cells workbook | log error when named range not found using Aspose.Cells for .NET | handling null named range Aspose.Cells C# example
// Tags: named range existence check Aspose.Cells | exception handling Aspose.Cells workbook | console logging Aspose.Cells errors | C# Aspose.Cells load workbook safely | Aspose.Cells save workbook with error handling

using System;
using System.IO;
using Aspose.Cells;

// // Loads an Excel file with Aspose.Cells, safely attempts to retrieve a named range called 'MyMissingRange' using a null check and try‑catch, logs appropriate messages to the console for missing or error conditions, and finally saves the workbook with error handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        Workbook workbook = null;

        // Load workbook safely
        try
        {
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        try
        {
            // Retrieve the named range; may be null if it doesn't exist
            Name namedRange = workbook.Worksheets.Names["MyMissingRange"];

            if (namedRange != null)
            {
                // Output the reference of the named range
                Console.WriteLine($"Named range found: refers to {namedRange.RefersTo}");
            }
            else
            {
                Console.WriteLine("Named range 'MyMissingRange' does not exist.");
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors while accessing the named range
            Console.WriteLine($"Error accessing named range: {ex.Message}");
        }

        // Save the workbook (optional)
        try
        {
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
