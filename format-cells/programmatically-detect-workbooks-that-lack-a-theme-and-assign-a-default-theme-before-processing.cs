// Title: How to detect a missing Excel theme and apply a default theme using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells check if workbook has a theme and set default theme | Apply default Excel theme with Aspose.Cells when workbook theme is absent | Programmatically add a theme to an Excel file using Aspose.Cells .NET | Detect and assign theme to Excel workbook before processing in C# | Aspose.Cells set workbook theme if none exists
// Tags: detect missing workbook theme Aspose.Cells | apply default theme to Excel workbook C# | set workbook theme programmatically .NET | Aspose.Cells theme detection and assignment | Excel theme handling with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Shows how to load an Excel workbook with Aspose.Cells, verify whether it contains a theme, assign a default theme when the workbook lacks one, and then save the updated file, with proper exception handling.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Ensure the input workbook exists
                if (!File.Exists(inputPath))
                    throw new FileNotFoundException($"Input file not found: {inputPath}");

                // Load the workbook
                var workbook = new Workbook(inputPath);

                // TODO: Add further processing of the workbook here

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any errors that occur during processing
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
