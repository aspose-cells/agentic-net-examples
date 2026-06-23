using System;
using System.IO;
using Aspose.Cells;

namespace WorksheetValidationProtectionDemoApp
{
    class WorksheetValidationProtectionDemo
    {
        static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Protect the worksheet with a password
            sheet.Protect(ProtectionType.All, "pwd123", null);

            // Unprotect the worksheet using the password
            sheet.Unprotect("pwd123");

            // Define the cell area for validation (cell B2)
            CellArea area = CellArea.CreateCellArea(1, 1, 1, 1); // Row 1, Column 1 (zero‑based)

            // Add a new validation to the collection for the defined area
            int validationIndex = sheet.Validations.Add(area);
            Validation validation = sheet.Validations[validationIndex];

            // Configure the validation (list type with a dropdown)
            validation.Type = ValidationType.List;
            validation.Formula1 = "OptionA,OptionB,OptionC";
            validation.InCellDropDown = true;
            validation.ShowInput = true;
            validation.InputMessage = "Select an option from the list.";
            validation.ErrorMessage = "Invalid selection.";

            // Re‑protect the worksheet with the same password
            sheet.Protect(ProtectionType.All, "pwd123", null);

            // Verify that the validation still exists after protection
            Validation check = sheet.Validations.GetValidationInCell(1, 1);
            Console.WriteLine("Validation type at B2: " + (check != null ? check.Type.ToString() : "None"));

            // Define output file path
            string outputPath = "WorksheetValidationProtectionDemo.xlsx";

            // Ensure we can write to the target location
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
    }
}