using System;
using Aspose.Cells;

namespace WorkbookValidationDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the workbook to be validated
            string inputPath = "SampleWorkbook.xlsx";

            // Load options – enable data validation checking while loading
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            loadOptions.CheckDataValid = true;

            // Load the workbook using the provided constructor (string, LoadOptions)
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Locate the worksheet named "Introduction"
            Worksheet introWorksheet = null;
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (ws.Name.Equals("Introduction", StringComparison.OrdinalIgnoreCase))
                {
                    introWorksheet = ws;
                    break;
                }
            }

            if (introWorksheet == null)
            {
                Console.WriteLine("Worksheet 'Introduction' not found.");
                return;
            }

            Console.WriteLine("Worksheet 'Introduction' found.");

            // Example validation: check specific cell values
            // Expected values can be adjusted as needed
            string expectedTitle = "Project Overview";
            string actualTitle = introWorksheet.Cells["A1"].StringValue;

            if (actualTitle == expectedTitle)
                Console.WriteLine("Cell A1 title is correct.");
            else
                Console.WriteLine($"Cell A1 title mismatch. Expected: '{expectedTitle}', Actual: '{actualTitle}'");

            // Example validation: numeric value in B2 should be between 0 and 100
            double? b2Value = introWorksheet.Cells["B2"].Value as double?;
            if (b2Value.HasValue && b2Value.Value >= 0 && b2Value.Value <= 100)
                Console.WriteLine("Cell B2 numeric value is within the expected range.");
            else
                Console.WriteLine("Cell B2 numeric value is out of range or not a number.");

            // Verify worksheet protection status and password (if protected)
            if (introWorksheet.IsProtected)
            {
                Console.WriteLine("Worksheet is protected. Verifying password...");

                // Replace "yourPassword" with the actual password used for protection
                bool passwordValid = introWorksheet.Protection.VerifyPassword("yourPassword");

                Console.WriteLine(passwordValid
                    ? "Password verification succeeded."
                    : "Password verification failed.");
            }
            else
            {
                Console.WriteLine("Worksheet is not protected.");
            }

            // Check for data validation rules on a specific cell (e.g., C3)
            Validation validation = introWorksheet.Validations.GetValidationInCell(2, 2); // Row 2, Column 2 => C3
            if (validation != null)
            {
                Console.WriteLine($"Data validation found on C3. Type: {validation.Type}");

                // Example: if the validation is a list, output the allowed values
                if (validation.Type == ValidationType.List && !string.IsNullOrEmpty(validation.Formula1))
                {
                    Console.WriteLine($"Allowed list values: {validation.Formula1}");
                }
            }
            else
            {
                Console.WriteLine("No data validation found on C3.");
            }

            // Save a copy of the workbook after validation (optional)
            string outputPath = "ValidatedWorkbook.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}