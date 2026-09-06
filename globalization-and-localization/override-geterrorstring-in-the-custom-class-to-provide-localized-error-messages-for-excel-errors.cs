// Title: Implement a custom ErrorStringProvider in Aspose.Cells to override GetErrorString for localized Excel error messages (C#)
// AI Prompts: Write a C# class that inherits from Aspose.Cells.ErrorStringProvider, overrides GetErrorString to return translated strings for Excel error codes, and registers it with a Workbook before formula calculation. | Show how to map specific Excel error enums (e.g., ErrorValueType.Div0) to custom messages in French and apply the provider to an existing workbook. | Demonstrate using the custom ErrorStringProvider to display localized error text when saving a workbook that contains formula errors.
// Common Searches: c# Aspose.Cells custom error string provider example for localization | override GetErrorString to translate #DIV/0! error in Aspose.Cells | how to display Excel formula errors in Spanish using Aspose.Cells | register custom ErrorStringProvider with Aspose.Cells before CalculateFormula | localized error messages for Excel formulas in Aspose.Cells C# tutorial
// Tags: ErrorStringProvider subclass Aspose.Cells C# | localize Excel error strings Aspose.Cells | custom GetErrorString logic for workbook | register error string provider before calculation | Excel formula error localization C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // This example demonstrates how to create a subclass of Aspose.Cells.ErrorStringProvider, override the GetErrorString method to return translated messages for Excel error codes (e.g., #DIV/0!, #VALUE!), register the provider with a Workbook, and then recalculate formulas so the localized error strings appear in the saved file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Example: set a formula that will generate a DIV/0 error.
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue(0);
                sheet.Cells["A2"].Formula = "=1/A1";

                // Recalculate formulas to populate the error.
                workbook.CalculateFormula();

                // Optionally load an existing workbook if the file exists.
                string inputPath = "input.xlsx";
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }

                // Save the workbook.
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
