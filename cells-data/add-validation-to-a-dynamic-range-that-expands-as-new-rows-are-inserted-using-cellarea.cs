// Title: Add whole-number data validation to a dynamically expanding column B range using CellArea in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add data validation that automatically includes newly inserted rows in column B | How to use CellArea to create a dynamic validation range in Aspose.Cells .NET | C# Aspose.Cells whole number validation for a column that grows when rows are added | Set up expandable data validation range with CellArea in an Aspose.Cells workbook | Apply between operator validation to a dynamic column range using Aspose.Cells API
// Tags: CellArea dynamic validation range Aspose.Cells | whole number between validation Aspose.Cells C# | expandable column B validation Aspose.Cells | add data validation to worksheet Aspose.Cells | C# Aspose.Cells validation for inserted rows

using Aspose.Cells;
using System;

// Creates a new workbook, defines a CellArea covering column B from row 1 to row 1000, adds a whole-number (1‑100) validation rule that applies to the dynamic range, and saves the file as DynamicValidation.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Define a dynamic range using CellArea.
        // This range starts at row 1 (index 0), column B (index 1) and extends to a large row number.
        // When new rows are inserted within this column, they fall inside the defined area.
        CellArea dynamicArea = new CellArea
        {
            StartRow = 0,      // Row 1
            StartColumn = 1,   // Column B
            EndRow = 999,      // Row 1000 (adjust as needed for expected data size)
            EndColumn = 1      // Column B
        };

        // Add a data validation rule to the dynamic range.
        // Example: whole number between 1 and 100.
        int validationIndex = sheet.Validations.Add(dynamicArea);
        Validation validation = sheet.Validations[validationIndex];
        validation.Type = ValidationType.WholeNumber;
        validation.Operator = OperatorType.Between;
        validation.Formula1 = "1";
        validation.Formula2 = "100";
        validation.ShowError = true;
        validation.ErrorTitle = "Invalid Input";
        validation.ErrorMessage = "Please enter a whole number between 1 and 100.";

        // Save the workbook
        workbook.Save("DynamicValidation.xlsx");
    }
}
