// Title: Create a time‑range data validation (09:00‑17:00) for the entire column T using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells set time validation for a whole column in Excel | How to restrict Excel column T to business hours (09:00 to 17:00) with Aspose.Cells | Aspose.Cells add between operator time validation programmatically in .NET | Create data validation for time values in column T using Aspose.Cells API | Set allowed time range 09:00‑17:00 for cells in column T with C#
// Tags: aspocells time range validation c# | excel column t validation aspocells | between operator validation aspocells | c# workbook business hours restriction | set allowed time values aspocells

using Aspose.Cells;
using System;

// The example creates a new workbook, defines a CellArea that covers the entire column T, adds a Time validation with the Between operator, limits acceptable entries to 09:00‑17:00, configures an error title and message, and saves the file as TimeValidation.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range for the validation (entire column T)
            int firstRow = 0;
            int lastRow = 1048575; // Excel's maximum row index (zero‑based)
            int columnIndex = 19;  // Column T (zero‑based)

            // Create a CellArea representing the target range
            CellArea area = new CellArea
            {
                StartRow = firstRow,
                StartColumn = columnIndex,
                EndRow = lastRow,
                EndColumn = columnIndex
            };

            // Add a new validation rule for the specified area (returns the index)
            int validationIndex = sheet.Validations.Add(area);
            Validation validation = sheet.Validations[validationIndex];

            // Set validation type to Time and operator to Between
            validation.Type = ValidationType.Time;
            validation.Operator = OperatorType.Between;

            // Define the allowed time range: 09:00 to 17:00
            validation.Formula1 = "09:00";
            validation.Formula2 = "17:00";

            // Optional: configure error message
            validation.ShowError = true;
            validation.ErrorTitle = "Invalid Time";
            validation.ErrorMessage = "Please enter a time between 09:00 and 17:00.";

            // Save the workbook
            workbook.Save("TimeValidation.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
