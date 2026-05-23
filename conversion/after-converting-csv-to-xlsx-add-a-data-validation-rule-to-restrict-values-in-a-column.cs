using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsDataValidationDemo
{
    class Program
    {
        static void Main()
        {
            // Paths for source CSV and intermediate XLSX file
            string csvPath = "input.csv";
            string xlsxPath = "output.xlsx";

            // -------------------------------------------------
            // 1. Convert CSV to XLSX using the provided utility
            // -------------------------------------------------
            // The ConversionUtility.Convert method is the approved way to perform the conversion.
            ConversionUtility.Convert(csvPath, xlsxPath);

            // -------------------------------------------------
            // 2. Load the generated workbook
            // -------------------------------------------------
            Workbook workbook = new Workbook(xlsxPath);

            // -------------------------------------------------
            // 3. Add a data‑validation rule to restrict values
            //    in column B (index 1) to whole numbers between 1 and 100
            // -------------------------------------------------
            Worksheet sheet = workbook.Worksheets[0];
            ValidationCollection validations = sheet.Validations;

            // Create a new validation object
            int validationIndex = validations.Add();
            Validation validation = validations[validationIndex];

            // Set validation type and criteria
            validation.Type = ValidationType.WholeNumber;
            validation.Operator = OperatorType.Between;
            validation.Formula1 = "1";   // Minimum allowed value
            validation.Formula2 = "100"; // Maximum allowed value

            // Apply the validation to the whole column B (rows 0‑1048575)
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 1,   // Column B
                EndRow = sheet.Cells.MaxDataRow, // Use existing max row or a large number
                EndColumn = 1
            };
            validation.AddArea(area);

            // -------------------------------------------------
            // 4. Save the workbook with the validation rule applied
            // -------------------------------------------------
            workbook.Save(xlsxPath, SaveFormat.Xlsx);

            Console.WriteLine("CSV converted to XLSX and data validation added successfully.");
        }
    }
}