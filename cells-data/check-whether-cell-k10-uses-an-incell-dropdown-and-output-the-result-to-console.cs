using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsInCellDropdownCheck
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: File \"{inputPath}\" not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve validation applied to cell K10
                Validation validation = worksheet.Cells["K10"].GetValidation();

                // Check if validation exists before accessing its properties
                bool hasInCellDropdown = validation != null && validation.InCellDropDown;

                Console.WriteLine($"Cell K10 in‑cell dropdown enabled: {hasInCellDropdown}");
            }
            catch (Exception ex)
            {
                // Log unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}