using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Configure load options to enforce data‑validation checking when the workbook is loaded.
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        loadOptions.CheckDataValid = true;

        // Create a new workbook (or load a template workbook if needed) using the load options.
        Workbook workbook = new Workbook(loadOptions);

        // Import XML data into the first worksheet starting at cell A1.
        // Adjust the file path, sheet name, and start cell as required.
        string xmlFilePath = "data.xml";
        workbook.ImportXml(xmlFilePath, "Sheet1", 0, 0);

        // Reference the worksheet that received the XML data.
        Worksheet sheet = workbook.Worksheets["Sheet1"];

        // ------------------------------------------------------------
        // Validation 1: Column A must contain whole numbers between 1 and 100.
        // ------------------------------------------------------------
        ValidationCollection validations = sheet.Validations;

        // Define the range for column A (rows 0‑1000, column index 0).
        CellArea columnARange = CellArea.CreateCellArea(0, 0, 1000, 0);
        int validationIndexA = validations.Add(columnARange);
        Validation validationA = validations[validationIndexA];

        validationA.Type = ValidationType.WholeNumber;
        validationA.Operator = OperatorType.Between;
        validationA.Formula1 = "1";
        validationA.Formula2 = "100";
        validationA.InputMessage = "Enter a whole number between 1 and 100.";
        validationA.ErrorMessage = "Value must be between 1 and 100.";
        validationA.ShowInput = true;
        validationA.ShowError = true;

        // ------------------------------------------------------------
        // Validation 2: Column B must be one of the predefined categories.
        // ------------------------------------------------------------
        // Define the range for column B (rows 0‑1000, column index 1).
        CellArea columnBRange = CellArea.CreateCellArea(0, 1, 1000, 1);
        int validationIndexB = validations.Add(columnBRange);
        Validation validationB = validations[validationIndexB];

        validationB.Type = ValidationType.List;
        validationB.InCellDropDown = true;
        // List values must be quoted and comma‑separated.
        validationB.Formula1 = "\"Electronics,Clothing,Food\"";
        validationB.InputMessage = "Select a category.";
        validationB.ErrorMessage = "Invalid category selected.";
        validationB.ShowInput = true;
        validationB.ShowError = true;

        // ------------------------------------------------------------
        // Save the workbook with the applied validations.
        // ------------------------------------------------------------
        workbook.Save("ValidatedOutput.xlsx");
    }
}