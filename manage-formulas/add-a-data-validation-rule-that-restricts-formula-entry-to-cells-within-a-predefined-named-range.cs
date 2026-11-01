// Title: Create list‑type data validation that references a named range in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add list validation using a named range | how to restrict cell input to a named range with Aspose.Cells for .NET | C# example of data validation formula referencing a named range in Excel | apply list‑type validation to A1:A10 using Aspose.Cells and a named range | Aspose.Cells validation Formula1 =MyRange C# code sample
// Tags: list validation with named range Aspose.Cells | data validation Formula1 named range C# | apply validation to cell area Aspose.Cells | restrict formula entry to predefined range Excel

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a new workbook, defines the cell range A1:A10, adds a list‑type validation that points to the named range "MyRange", sets an error title and message, and saves the workbook as Output.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (or load an existing one if needed)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Define the target cells for validation (A1:A10)
                CellArea validationArea = new CellArea
                {
                    StartRow = 0,   // Row 1 (zero‑based)
                    EndRow = 9,     // Row 10
                    StartColumn = 0,// Column A
                    EndColumn = 0   // Column A
                };

                // Add a list‑type validation referencing a named range "MyRange"
                int validationIndex = worksheet.Validations.Add(validationArea);
                Validation validation = worksheet.Validations[validationIndex];
                validation.Type = ValidationType.List;
                validation.Formula1 = "=MyRange";
                validation.ShowError = true;
                validation.ErrorTitle = "Invalid Entry";
                validation.ErrorMessage = "Please select a value from the predefined list.";

                // Save the workbook
                string outputPath = "Output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
