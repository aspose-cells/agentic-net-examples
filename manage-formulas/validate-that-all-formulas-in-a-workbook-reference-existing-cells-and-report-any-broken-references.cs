// Title: Identify and list broken #REF! formula references in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write a C# console program with Aspose.Cells that opens a given .xlsx file, scans every worksheet for formulas containing "#REF!", and prints the sheet name, cell address, and full formula for each broken reference. | Create a reusable C# method that receives a Workbook object and returns a collection of cells whose formulas have invalid references, leveraging the Aspose.Cells API. | Add robust error handling to a C# Aspose.Cells script that verifies the input file exists before loading, then logs any detected #REF! errors to the console.
// Common Searches: C# Aspose.Cells how to detect #REF! errors in all worksheets of an Excel file | list cells with broken formula references using Aspose.Cells .NET | validate formulas for missing references in an .xlsx workbook with Aspose.Cells | Aspose.Cells iterate through cells to find invalid #REF! formulas in C# console app
// Tags: detect broken #REF! formulas Aspose.Cells | iterate worksheets cells Aspose.Cells C# | validate Excel formulas .NET | log invalid formula references Aspose.Cells | check formula integrity Excel .xlsx Aspose

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // A C# console example that loads an .xlsx workbook with Aspose.Cells, iterates through all worksheets and cells, identifies formulas containing the "#REF!" token, and outputs the sheet name, cell address, and formula for each broken reference, with file‑existence checks and exception handling.
    class FormulaReferenceValidator
    {
        static void Main()
        {
            try
            {
                // Path to the input workbook
                string inputPath = "input.xlsx";

                // Verify that the file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                bool hasBrokenReferences = false;

                // Iterate through all worksheets and cells to detect broken references (#REF!)
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;
                    foreach (Cell cell in cells)
                    {
                        if (cell.IsFormula)
                        {
                            string formula = cell.Formula;
                            if (!string.IsNullOrEmpty(formula) && formula.Contains("#REF!"))
                            {
                                hasBrokenReferences = true;
                                Console.WriteLine("Broken reference found:");
                                Console.WriteLine($"  Sheet : {sheet.Name}");
                                Console.WriteLine($"  Cell  : {cell.Name}");
                                Console.WriteLine($"  Formula: {formula}");
                                Console.WriteLine();
                            }
                        }
                    }
                }

                if (!hasBrokenReferences)
                {
                    Console.WriteLine("No broken references were detected in any formula.");
                }

                // Optionally, save a copy of the workbook (if modifications were made)
                // workbook.Save("output.xlsx");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
