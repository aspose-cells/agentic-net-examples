using System;
using System.IO;
using Aspose.Cells;

class ValidationErrorReporter
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // -------------------------------------------------
        // Add sample data validations to demonstrate reporting
        // -------------------------------------------------

        // Validation for cell A1: Whole number between 10 and 20
        CellArea areaA1 = new CellArea { StartRow = 0, StartColumn = 0, EndRow = 0, EndColumn = 0 };
        int idxA1 = sheet.Validations.Add(areaA1);
        Validation valA1 = sheet.Validations[idxA1];
        valA1.Type = ValidationType.WholeNumber;
        valA1.Operator = OperatorType.Between;
        valA1.Formula1 = "10";
        valA1.Formula2 = "20";
        valA1.ErrorTitle = "Invalid Input";
        valA1.ErrorMessage = "Value must be between 10 and 20";
        valA1.ShowError = true;
        valA1.AlertStyle = ValidationAlertType.Stop;

        // Validation for cell B1: List of allowed values
        CellArea areaB1 = new CellArea { StartRow = 0, StartColumn = 1, EndRow = 0, EndColumn = 1 };
        int idxB1 = sheet.Validations.Add(areaB1);
        Validation valB1 = sheet.Validations[idxB1];
        valB1.Type = ValidationType.List;
        valB1.Formula1 = "\"Red,Green,Blue\"";
        valB1.ErrorTitle = "Invalid Color";
        valB1.ErrorMessage = "Select a color from the list.";
        valB1.ShowError = true;
        valB1.AlertStyle = ValidationAlertType.Information;

        // -------------------------------------------------
        // Write validation error details to a text file
        // -------------------------------------------------
        string reportPath = "ValidationErrors.txt";

        using (StreamWriter writer = new StreamWriter(reportPath))
        {
            writer.WriteLine("Validation Error Report");
            writer.WriteLine("========================");
            writer.WriteLine();

            // Iterate through all validations in the worksheet
            for (int i = 0; i < sheet.Validations.Count; i++)
            {
                Validation v = sheet.Validations[i];
                writer.WriteLine($"Validation #{i + 1}");
                writer.WriteLine($"Error Title   : {v.ErrorTitle}");
                writer.WriteLine($"Error Message : {v.ErrorMessage}");
                writer.WriteLine($"Show Error    : {v.ShowError}");
                writer.WriteLine($"Alert Style   : {v.AlertStyle}");
                writer.WriteLine();
            }
        }

        // -------------------------------------------------
        // Save the workbook (optional, demonstrates normal lifecycle)
        // -------------------------------------------------
        workbook.Save("ValidationDemo.xlsx");
    }
}