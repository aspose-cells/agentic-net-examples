// Title: C# – Add List Data Validation to Column B and Freeze Panes with Aspose.Cells
// Description: Creates a new workbook, defines a CellArea for column B (rows 0‑100), adds a list‑type validation with a dropdown and custom input/error messages, freezes the first two columns so the validated column stays visible while scrolling, and saves the file as an XLSX workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# data validation | list validation | dropdown validation | freeze panes | column B | Excel workbook | CellArea | validation messages | SaveFormat.Xlsx | Aspose.Cells example | GitHub sample
// Common Searches: Aspose.Cells add list validation to a column | How to freeze columns in Aspose.Cells .NET | C# example for data validation dropdown with Aspose.Cells | Keep validation column visible by freezing panes Aspose.Cells | Aspose.Cells set input and error messages for validation
// Developer Intent: Generate an Excel file, apply a list‑type data validation to column B, and freeze the first two columns so the validation dropdown remains on screen.
// Use Cases: Design a data‑entry template where users must pick from predefined options in column B while scrolling horizontally. | Create a reporting sheet that locks a validated status column in place to prevent it from scrolling out of view during large data entry. | Export a formatted workbook with a validated category column that stays visible when the worksheet is scrolled.
// AI Prompts: Write C# code using Aspose.Cells to add a list validation to column C rows 1‑200 with custom input and error messages, then freeze the first three columns. | Provide an Aspose.Cells .NET example that applies a numeric range validation to column D and keeps the column visible by freezing panes. | Show how to set up a dropdown validation with three options in column B and freeze the first two columns so the validation column stays on screen.

using Aspose.Cells;

// Creates a new workbook, defines a CellArea for column B (rows 0‑100), adds a list‑type validation with a dropdown and custom input/error messages, freezes the first two columns so the validated column stays visible while scrolling, and saves the file as an XLSX workbook using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Define the validation area for column B (index 1) from row 0 to row 100
        CellArea validationArea = new CellArea
        {
            StartRow = 0,
            EndRow = 100,
            StartColumn = 1,
            EndColumn = 1
        };

        // Add a list‑type validation to the defined area
        int validationIndex = worksheet.Validations.Add(validationArea);
        Validation validation = worksheet.Validations[validationIndex];
        validation.Type = ValidationType.List;
        validation.InCellDropDown = true;
        validation.Formula1 = "Option1,Option2,Option3";

        // Optional UI messages
        validation.ShowInput = true;
        validation.InputTitle = "Select Option";
        validation.InputMessage = "Choose one of the listed options.";
        validation.ShowError = true;
        validation.ErrorTitle = "Invalid Selection";
        validation.ErrorMessage = "Please select a valid option from the list.";

        // Freeze the first two columns so column B (the validated column) stays visible
        worksheet.FreezePanes(0, 2, 0, 2);

        // Save the workbook
        workbook.Save("ColumnValidationAndFreeze.xlsx", SaveFormat.Xlsx);
    }
}
