// Title: How to read a cell’s data‑validation type and formulas (Formula1, Formula2) with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells get data validation type of a specific cell | read validation Formula1 and Formula2 from cell J5 using Aspose.Cells | Aspose.Cells retrieve data validation details for an Excel cell in .NET | how to check if a worksheet cell has data validation with Aspose.Cells C# | extract data validation rules from a worksheet cell using Aspose.Cells API
// Tags: Aspose.Cells read cell validation .NET | data validation extraction Aspose.Cells C# | cell validation type retrieval Excel .NET | GetValidation API usage Aspose.Cells | read validation formulas Excel workbook C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing XLSX workbook, accesses the first worksheet, obtains cell J5, and uses the GetValidation method to fetch any data‑validation applied. It then prints the validation type along with Formula1 and Formula2, handling cases where no validation exists and catching CellsException for password‑protected or corrupted files.
    class Program
    {
        static void Main()
        {
            string inputPath = "input.xlsx";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"File not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook (no password). If the file is password‑protected,
                // a CellsException will be thrown and handled in the catch block below.
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Get cell J5
                Cell cell = worksheet.Cells["J5"];

                // Retrieve any data validation applied to the cell
                Validation validation = cell.GetValidation();

                if (validation != null)
                {
                    Console.WriteLine($"Validation Type: {validation.Type}");
                    Console.WriteLine($"Formula1: {validation.Formula1}");
                    Console.WriteLine($"Formula2: {validation.Formula2}");
                }
                else
                {
                    Console.WriteLine("Cell J5 does not have any data validation applied.");
                }

                // (Optional) Save the workbook if any changes were made
                // workbook.Save("output.xlsx");
            }
            catch (CellsException ce)
            {
                // Handle password‑protected workbook or other CellsException cases
                Console.WriteLine($"CellsException: {ce.Message}");

                // Example of loading a password‑protected workbook:
                // string password = "yourPassword";
                // LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx) { Password = password };
                // Workbook workbook = new Workbook(inputPath, loadOptions);
            }
            catch (Exception ex)
            {
                // General exception handling
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
