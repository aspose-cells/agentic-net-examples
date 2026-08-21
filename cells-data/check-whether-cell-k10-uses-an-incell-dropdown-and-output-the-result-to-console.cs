// Title: Check if cell K10 has an in‑cell dropdown using Aspose.Cells for .NET
// Description: Loads an Excel workbook, accesses the first worksheet, reads cell K10, examines its Validation object, and prints whether the Validation.InCellDropDown flag is true. Includes error handling for missing files and absent validation.
// Keywords: Aspose.Cells | C# | in‑cell dropdown | data validation | Validation.InCellDropDown | Excel | .NET | check dropdown | cell K10 | workbook loading
// Common Searches: Aspose.Cells detect dropdown in cell | C# check if Excel cell has data validation list | How to read Validation.InCellDropDown property | Determine if cell K10 contains a dropdown using Aspose | Read cell validation with Aspose.Cells .NET
// Developer Intent: Identify whether cell K10 contains an in‑cell dropdown and output the result.
// Use Cases: Validate that required dropdowns exist before processing user‑entered data. | Create an inventory of cells that use data‑validation lists for documentation or migration. | Execute conditional business logic only when a specific cell is configured with a dropdown. | Verify template integrity prior to data entry in automated workflows.
// AI Prompts: Write a reusable method that takes a worksheet and cell address and returns true if the cell has an in‑cell dropdown using Aspose.Cells. | Generate code to scan an entire worksheet and list all cells that contain in‑cell dropdowns. | Add detailed logging to the dropdown‑check program, recording cell address, validation type, and dropdown status. | Convert the console example into an async service method suitable for ASP.NET Core applications.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsInCellDropdownCheck
{
    // Loads an Excel workbook, accesses the first worksheet, reads cell K10, examines its Validation object, and prints whether the Validation.InCellDropDown flag is true. Includes error handling for missing files and absent validation.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: File not found – {inputPath}");
                return;
            }

            Workbook workbook;
            try
            {
                // Load the workbook
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
                return;
            }

            try
            {
                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Get cell K10
                Cell cell = worksheet.Cells["K10"];

                // Retrieve validation applied to the cell (null if none)
                Validation validation = cell.GetValidation();

                // Determine whether an in‑cell dropdown is enabled
                bool hasInCellDropdown = false;
                if (validation != null)
                {
                    hasInCellDropdown = validation.InCellDropDown;
                }

                // Output the result
                Console.WriteLine($"Cell K10 uses in‑cell dropdown: {hasInCellDropdown}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Runtime error: {ex.Message}");
            }
        }
    }
}
