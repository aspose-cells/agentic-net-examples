using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    public class EnableInCellDropdown
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Define the validation area for column H (zero‑based column index 7)
                // Apply the validation to rows 1‑100 (indices 0‑99)
                CellArea area = new CellArea
                {
                    StartRow = 0,
                    EndRow = 99,
                    StartColumn = 7,
                    EndColumn = 7
                };

                // Add a new validation to the worksheet's validation collection
                ValidationCollection validations = worksheet.Validations;
                int validationIndex = validations.Add(area);
                Validation validation = validations[validationIndex];

                // Set the validation type to List and provide the list values
                validation.Type = ValidationType.List;
                validation.Formula1 = "Option1,Option2,Option3";

                // Enable the in‑cell dropdown list
                validation.InCellDropDown = true;

                // Save the workbook
                string outputPath = "ColumnH_InCellDropdown.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            EnableInCellDropdown.Run();
        }
    }
}