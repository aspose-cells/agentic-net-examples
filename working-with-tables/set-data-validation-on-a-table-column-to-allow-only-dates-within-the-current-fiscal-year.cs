// Title: How to add fiscal‑year date validation to a column in an Excel table using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set date validation for a ListObject column based on current fiscal year | C# add Excel table column validation to allow only dates from July to June using Aspose.Cells | How to calculate fiscal year range and apply it as data validation in an Aspose.Cells workbook | Apply date validation to a specific column of an Excel table with Aspose.Cells .NET | Aspose.Cells validation for dynamic date range in a table column C# example
// Tags: Aspose.Cells ListObject validation | C# fiscal year date range calculation | Excel table column validation Aspose.Cells | CellArea definition for data validation | Dynamic date validation in .NET Excel

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example loads an existing workbook, locates the first worksheet and its first ListObject (Excel table), computes the current fiscal year (July 1 to June 30), defines the data rows of a chosen table column, and adds a date validation rule that restricts entries to dates within that fiscal year. The validation displays an error message for out‑of‑range values, and the updated workbook is saved to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one table (ListObject)
            if (worksheet.ListObjects.Count == 0)
                throw new InvalidOperationException("No tables (ListObjects) found on the worksheet.");

            // Get the first table
            ListObject table = worksheet.ListObjects[0];

            // Column index within the table (zero‑based). Change as needed.
            int tableColumnIndex = 1;

            // Determine the data range of the column (excluding the header row)
            int firstDataRow = table.DataRange.FirstRow;
            int lastDataRow = table.DataRange.FirstRow + table.DataRange.RowCount - 1;
            // Use DataRange.FirstColumn to get the absolute column index of the table's first column
            int columnAbsoluteIndex = table.DataRange.FirstColumn + tableColumnIndex;

            // Define the cell area for the validation
            CellArea validationArea = new CellArea
            {
                StartRow = firstDataRow,
                EndRow = lastDataRow,
                StartColumn = columnAbsoluteIndex,
                EndColumn = columnAbsoluteIndex
            };

            // Calculate fiscal year (July 1 to June 30)
            DateTime today = DateTime.Today;
            int fiscalYearStartYear = today.Month >= 7 ? today.Year : today.Year - 1;
            DateTime fiscalYearStart = new DateTime(fiscalYearStartYear, 7, 1);
            DateTime fiscalYearEnd = fiscalYearStart.AddYears(1).AddDays(-1);

            // Add date validation for the defined area
            ValidationCollection validations = worksheet.Validations;
            int validationIndex = validations.Add(validationArea);
            Validation validation = validations[validationIndex];
            validation.Type = ValidationType.Date;
            validation.Operator = OperatorType.Between;
            validation.Formula1 = fiscalYearStart.ToOADate().ToString(); // start date
            validation.Formula2 = fiscalYearEnd.ToOADate().ToString();   // end date
            validation.ShowError = true;
            validation.ErrorTitle = "Invalid Date";
            validation.ErrorMessage = $"Please enter a date between {fiscalYearStart:d} and {fiscalYearEnd:d}.";

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
