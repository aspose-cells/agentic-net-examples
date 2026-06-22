using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ValidationFromHiddenSheet
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first (visible) worksheet
                Worksheet visibleSheet = workbook.Worksheets[0];
                visibleSheet.Name = "DataEntry";

                // Add a hidden worksheet that will hold the allowed list values
                Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenValues");
                hiddenSheet.IsVisible = false; // Hide the sheet

                // Populate the hidden sheet with the list items (A1:A5)
                string[] allowedValues = { "Apple", "Banana", "Cherry", "Date", "Elderberry" };
                for (int i = 0; i < allowedValues.Length; i++)
                {
                    hiddenSheet.Cells[i, 0].PutValue(allowedValues[i]); // Column A (index 0)
                }

                // Define the cell (B1) on the visible sheet where the validation will be applied
                CellArea validationArea = CellArea.CreateCellArea(0, 1, 0, 1); // Row 0, Column 1 => B1

                // Add a new validation to the visible sheet
                int validationIndex = visibleSheet.Validations.Add(validationArea);
                Validation validation = visibleSheet.Validations[validationIndex];

                // Configure the validation as a List that references the hidden range
                validation.Type = ValidationType.List;
                validation.Formula1 = "HiddenValues!A1:A5"; // Reference to hidden sheet range
                validation.InCellDropDown = true; // Show the drop‑down arrow in the cell

                // Save the workbook
                string outputPath = "ValidationFromHiddenSheet.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            ValidationFromHiddenSheet.Run();
        }
    }
}