// Title: Enforce integer values between 10 and 500 in column G with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to restrict column G to integers between 10 and 500 | Add whole‑number data validation to an Excel column using Aspose.Cells .NET | Set numeric range validation for column G in a workbook with Aspose.Cells | Aspose.Cells example for integer range validation (10‑500) in C# | Validate Excel cells to accept only numbers 10‑500 using Aspose.Cells
// Tags: Aspose.Cells integer validation | C# Excel column G numeric range | data validation whole number Aspose.Cells | Excel column validation between 10 and 500 | Aspose.Cells set validation for column

using Aspose.Cells;
using System;
using System.IO;

// The code loads or creates an Excel workbook, selects the first worksheet, determines the last used row, defines a range covering column G, and adds a whole‑number validation that only allows integers from 10 to 500. The workbook is then saved with the validation applied.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook if it exists; otherwise create a new one.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook with one worksheet
            }

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the last used row; if none, default to row 0.
            int lastRow = sheet.Cells.MaxDataRow >= 0 ? sheet.Cells.MaxDataRow : 0;

            // Define the range covering the entire column G (zero‑based column index 6)
            CellArea range = new CellArea
            {
                StartRow = 0,
                EndRow = lastRow,
                StartColumn = 6,
                EndColumn = 6
            };

            // Add a validation rule that restricts values to whole numbers between 10 and 500
            int validationIndex = sheet.Validations.Add(range);
            Validation validation = sheet.Validations[validationIndex];
            validation.Type = ValidationType.WholeNumber;   // integer validation
            validation.Operator = OperatorType.Between;
            validation.Formula1 = "10";
            validation.Formula2 = "500";
            validation.ShowError = true;
            validation.ErrorTitle = "Invalid Input";
            validation.ErrorMessage = "Please enter an integer between 10 and 500.";

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
