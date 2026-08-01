// Title: C# – Convert CSV to XLSX and Add Whole‑Number Validation to Column B with Aspose.Cells
// Description: This example shows how to use Aspose.Cells for .NET to (1) convert a CSV file to an XLSX workbook via ConversionUtility, (2) define a CellArea covering column B, (3) create a Validation object that restricts entries to whole numbers between 1 and 100, and (4) save the validated workbook. The code demonstrates the required API calls and configuration of input and error messages.
// Keywords: Aspose.Cells | CSV to XLSX conversion | C# data validation | ValidationCollection | whole number range | column B validation | ConversionUtility | Excel template generation | numeric constraint | SaveFormat.Xlsx
// Common Searches: Aspose.Cells convert CSV to XLSX C# | add whole number validation column B Aspose.Cells | C# data validation range 1 to 100 Excel | how to use ValidationCollection in Aspose.Cells | restrict Excel column values with Aspose.Cells
// Developer Intent: The developer needs to transform a CSV file into an XLSX workbook and then enforce a numeric validation rule (1‑100) on every cell of column B.
// Use Cases: Generate an editable Excel report from a CSV export while guaranteeing that IDs entered in column B stay within a predefined numeric range. | Create a pre‑filled template where users can only input rating scores (1‑100) in column B, reducing data‑entry errors. | Automate the production of Excel files from CSV sources that must comply with downstream validation rules for numeric fields.
// AI Prompts: Write C# code using Aspose.Cells to convert a CSV file to XLSX and add a date‑range validation to column C (today to 30 days ahead). | Explain how ValidationCollection.Add works in Aspose.Cells and how to customize input and error messages for different validation types. | Create a unit test in C# that confirms the whole‑number validation on column B is active after converting a CSV with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsDataValidationExample
{
    // This example shows how to use Aspose.Cells for .NET to (1) convert a CSV file to an XLSX workbook via ConversionUtility, (2) define a CellArea covering column B, (3) create a Validation object that restricts entries to whole numbers between 1 and 100, and (4) save the validated workbook. The code demonstrates the required API calls and configuration of input and error messages.
    class Program
    {
        static void Main()
        {
            // Paths for source CSV and intermediate XLSX file
            string csvPath = "input.csv";
            string xlsxPath = "intermediate.xlsx";

            // -------------------------------------------------
            // 1. Convert CSV to XLSX using the provided rule
            // -------------------------------------------------
            // The ConversionUtility.Convert(string, string) method is the
            // mandated way to perform the conversion.
            ConversionUtility.Convert(csvPath, xlsxPath);

            // -------------------------------------------------
            // 2. Load the converted workbook (standard load)
            // -------------------------------------------------
            Workbook workbook = new Workbook(xlsxPath);

            // -------------------------------------------------
            // 3. Add a data validation rule to restrict values
            //    in column B (index 1) to whole numbers between 1 and 100.
            // -------------------------------------------------
            Worksheet sheet = workbook.Worksheets[0];
            ValidationCollection validations = sheet.Validations;

            // Define the area: entire column B (rows 0 to 65535)
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 65535,
                StartColumn = 1,
                EndColumn = 1
            };

            // Add a new validation and configure it
            int validationIndex = validations.Add(area);
            Validation validation = validations[validationIndex];
            validation.Type = ValidationType.WholeNumber;
            validation.Operator = OperatorType.Between;
            validation.Formula1 = "1";   // Minimum allowed value
            validation.Formula2 = "100"; // Maximum allowed value
            validation.ShowInput = true;
            validation.InputMessage = "Enter a whole number between 1 and 100.";
            validation.ShowError = true;
            validation.ErrorMessage = "Invalid entry. Value must be between 1 and 100.";
            validation.ErrorTitle = "Invalid Input";

            // -------------------------------------------------
            // 4. Save the workbook with the validation applied
            // -------------------------------------------------
            // Use the standard Workbook.Save method as required.
            workbook.Save(xlsxPath, SaveFormat.Xlsx);

            Console.WriteLine("Data validation added and workbook saved successfully.");
        }
    }
}
