// Title: Create a time validation for column T to allow only 09:00‑17:00 entries with Aspose.Cells in C#
// Common Searches: Aspose.Cells C# how to restrict column T to business hours 09:00 to 17:00 | C# set Excel time validation between 9 AM and 5 PM using Aspose.Cells | Add data validation for time range in Excel column with Aspose.Cells .NET | Configure ValidationType.Time with OADate boundaries in Aspose.Cells workbook | Excel time entry validation for specific column using Aspose.Cells API
// Tags: Aspose.Cells time validation C# | Excel column T data validation .NET | ValidationType.Time with OADate Aspose.Cells | business hours validation Aspose.Cells | C# set time range validation Excel

using System;
using Aspose.Cells;

// The example creates a workbook, defines the entire column T as a CellArea, adds a Validation of type Time with the Between operator, sets the lower and upper limits to 09:00 and 17:00 using OLE Automation dates, configures an error alert, and saves the file as ValidatedWorkbook.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Define the range covering the entire column T (index 19)
        CellArea timeRange = new CellArea
        {
            StartRow = 0,
            EndRow = sheet.Cells.MaxRow,   // 0‑based index, covers all rows
            StartColumn = 19,              // Column T
            EndColumn = 19
        };

        // Add a time validation to the defined range
        int validationIndex = sheet.Validations.Add(timeRange);
        Validation timeValidation = sheet.Validations[validationIndex];

        // Set validation type to Time and require a value between 09:00 and 17:00
        timeValidation.Type = ValidationType.Time;
        timeValidation.Operator = OperatorType.Between;

        // Convert the boundary times to OLE Automation dates (double) and assign as formulas
        double startTime = DateTime.Today.AddHours(9).ToOADate();   // 09:00
        double endTime   = DateTime.Today.AddHours(17).ToOADate();  // 17:00
        timeValidation.Formula1 = startTime.ToString();
        timeValidation.Formula2 = endTime.ToString();

        // Optional: display an error message when the rule is violated
        timeValidation.ShowError = true;
        timeValidation.ErrorTitle = "Invalid Time";
        timeValidation.ErrorMessage = "Please enter a time between 09:00 and 17:00.";

        // Save the workbook
        workbook.Save("ValidatedWorkbook.xlsx");
    }
}
