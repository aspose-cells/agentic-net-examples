// Title: C# – Add List Data Validation to Column C and Freeze It with Aspose.Cells
// Description: Creates a new workbook, applies a drop‑down list validation to column C (rows 0‑1000) using Aspose.Cells, freezes the column at cell D1 so the validation stays visible while scrolling, and saves the file as ColumnValidationAndFreeze.xlsx.
// Keywords: Aspose.Cells C# | Excel data validation list | list validation Aspose.Cells | freeze column Aspose.Cells | FreezePanes C# | CellArea validation | ValidationType.List | Excel automation .NET | drop‑down list Excel C# | Aspose.Cells example
// Common Searches: Aspose.Cells add drop‑down list to a column C# | How to freeze a column after adding validation with Aspose.Cells | C# code for data validation and freeze panes in Excel | Aspose.Cells freeze panes example | Create list validation in Excel using Aspose.Cells .NET
// Developer Intent: Generate a worksheet that contains a list‑type validation on column C and keeps that column fixed on screen by freezing panes.
// Use Cases: Apply a drop‑down list with predefined options to a specific column range. | Ensure the validated column remains visible during horizontal scrolling by freezing panes. | Combine data validation and frozen panes in a single Aspose.Cells workflow before saving the workbook.
// AI Prompts: Write C# Aspose.Cells code to add a list validation to column B and freeze column B. | Show how to configure Validation.InCellDropDown, InputTitle, and InputMessage, then freeze the column containing the validation. | Provide an Aspose.Cells example that saves an Excel file after applying both data validation and FreezePanes.

using Aspose.Cells;

// Creates a new workbook, applies a drop‑down list validation to column C (rows 0‑1000) using Aspose.Cells, freezes the column at cell D1 so the validation stays visible while scrolling, and saves the file as ColumnValidationAndFreeze.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Define a validation area that covers column C (index 2) from row 0 to row 1000
        CellArea validationArea = CellArea.CreateCellArea(0, 2, 1000, 2);

        // Add the validation to the worksheet's validation collection
        int validationIndex = sheet.Validations.Add(validationArea);
        Validation validation = sheet.Validations[validationIndex];

        // Configure the validation as a drop‑down list
        validation.Type = ValidationType.List;
        validation.Formula1 = "Option1,Option2,Option3";
        validation.InCellDropDown = true;
        validation.ShowInput = true;
        validation.InputTitle = "Select Option";
        validation.InputMessage = "Choose one of the listed options.";

        // Freeze columns up to column C so the validation column stays visible while scrolling
        // Freeze at cell D1 (column index 3) with 0 frozen rows and 1 frozen column
        sheet.FreezePanes("D1", 0, 1);

        // Save the workbook
        workbook.Save("ColumnValidationAndFreeze.xlsx", SaveFormat.Xlsx);
    }
}
