// Title: Load an Excel workbook in C# with Aspose.Cells preserving cell formatting and disabling formula evaluation using LoadOptions
// AI Prompts: Write C# code that uses Aspose.Cells LoadOptions to open an .xlsx file, keep all cell styles unchanged, and prevent any formulas from being calculated before saving. | Show how to verify that an Excel file exists, then load it with LoadOptions so formatting is retained and formulas remain unevaluated, and finally save the result to a new file. | Demonstrate using Aspose.Cells LoadOptions to load a workbook without triggering formula calculation, then export the workbook while preserving the original formatting.
// Common Searches: Aspose.Cells C# load workbook without evaluating formulas and keep original formatting | How to preserve cell styles when opening Excel file with LoadOptions in .NET | Load Excel file with Aspose.Cells while skipping formula calculation | C# example for loading .xlsx and saving with unchanged formatting using Aspose.Cells | Using LoadOptions to ignore formula results in Aspose.Cells workbook
// Tags: preserve formatting on workbook open Aspose.Cells | skip formula calculation with LoadOptions | export workbook retaining original styles .NET | verify Excel file existence before loading Aspose.Cells | load xlsx using LoadOptions C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // // This program checks for the presence of input.xlsx, loads it with LoadOptions (Xlsx format) which retains cell formatting and leaves formulas unevaluated, and then saves the result as output_preserved.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source workbook
            string sourcePath = "input.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Error: The file '{sourcePath}' was not found.");
                return;
            }

            try
            {
                // Configure LoadOptions (formulas are loaded but not evaluated by default)
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

                // Load the workbook with the specified options
                Workbook workbook = new Workbook(sourcePath, loadOptions);

                // Path for the output workbook
                string outputPath = "output_preserved.xlsx";

                try
                {
                    // Save the workbook; formatting is preserved and formulas remain unevaluated
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Error saving workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display the error message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
