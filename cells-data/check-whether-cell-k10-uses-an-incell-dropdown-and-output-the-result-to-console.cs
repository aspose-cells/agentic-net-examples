// Title: How to determine if cell K10 contains an in‑cell dropdown list using Aspose.Cells for .NET and display the result in the console
// AI Prompts: Generate a C# console program that loads an Excel file with Aspose.Cells, accesses cell K10, checks its Validation.InCellDropDown property, and writes the boolean outcome to the console. | Write C# code that accepts a cell address as input, uses Aspose.Cells to retrieve the cell's Validation object, and returns whether an in‑cell dropdown is enabled. | Create a reusable method in C# that takes a Worksheet and a cell reference, returns true if the cell has an in‑cell dropdown list via Aspose.Cells, and prints the result.
// Common Searches: Aspose.Cells C# check if a specific cell has an in‑cell dropdown list | How to read the InCellDropDown property of a cell validation in Aspose.Cells | C# console application to detect dropdown validation in Excel using Aspose.Cells | Aspose.Cells validation dropdown detection example | Determine whether Excel cell K10 contains a data validation list with Aspose.Cells
// Tags: Aspose.Cells GetValidation InCellDropDown | C# detect Excel cell dropdown list | Aspose.Cells read cell validation | Console output Excel validation result | Check specific cell dropdown Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsInCellDropdownCheck
{
    // The sample loads 'input.xlsx', accesses the first worksheet, retrieves cell K10, obtains its Validation object, evaluates the InCellDropDown flag, and writes the boolean result to the console while handling missing files and exceptions.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";

                // Verify that the input workbook exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (or specify the required sheet)
                Worksheet worksheet = workbook.Worksheets[0];

                // Get the cell K10
                Cell cell = worksheet.Cells["K10"];

                // Retrieve the validation applied to the cell (may be null)
                Validation validation = cell.GetValidation();

                // Determine whether the validation displays an in‑cell dropdown
                bool hasInCellDropdown = validation != null && validation.InCellDropDown;

                // Output the result to the console
                Console.WriteLine($"Cell K10 uses an in‑cell dropdown: {hasInCellDropdown}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
