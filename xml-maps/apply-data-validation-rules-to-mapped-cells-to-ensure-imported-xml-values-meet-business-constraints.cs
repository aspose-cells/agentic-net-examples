using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Enable data‑validation checking when the workbook is loaded or saved
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        loadOptions.CheckDataValid = true;

        // Create a new workbook (empty template)
        Workbook workbook = new Workbook();

        // Import XML data into the first worksheet starting at cell A1
        // Assume the XML file exists at the specified path
        workbook.ImportXml("data.xml", "Sheet1", 0, 0);

        // Get reference to the worksheet that now contains the imported data
        Worksheet sheet = workbook.Worksheets["Sheet1"];

        // ------------------------------------------------------------
        // Apply a decimal validation to the "Price" column (B) rows 1‑100
        // ------------------------------------------------------------
        ValidationCollection validations = sheet.Validations;

        // Define the area for the validation (B1:B100)
        CellArea priceArea = new CellArea
        {
            StartRow = 0,      // Row 1 (zero‑based)
            EndRow = 99,       // Row 100
            StartColumn = 1,   // Column B (zero‑based)
            EndColumn = 1
        };

        // Add the validation to the collection and configure it
        int priceValIndex = validations.Add(priceArea);
        Validation priceValidation = validations[priceValIndex];
        priceValidation.Type = ValidationType.Decimal;
        priceValidation.Operator = OperatorType.Between;
        priceValidation.Formula1 = "0";        // Minimum price
        priceValidation.Formula2 = "10000";    // Maximum price
        priceValidation.InputMessage = "Enter a price between 0 and 10,000.";
        priceValidation.ErrorMessage = "Invalid price value.";
        priceValidation.ShowInput = true;
        priceValidation.ShowError = true;

        // ------------------------------------------------------------
        // Apply a whole‑number validation to the "Quantity" column (C) rows 1‑100
        // ------------------------------------------------------------
        CellArea qtyArea = new CellArea
        {
            StartRow = 0,
            EndRow = 99,
            StartColumn = 2,   // Column C
            EndColumn = 2
        };

        int qtyValIndex = validations.Add(qtyArea);
        Validation qtyValidation = validations[qtyValIndex];
        qtyValidation.Type = ValidationType.WholeNumber;
        qtyValidation.Operator = OperatorType.Between;
        qtyValidation.Formula1 = "1";      // Minimum quantity
        qtyValidation.Formula2 = "1000";   // Maximum quantity
        qtyValidation.InputMessage = "Enter a quantity between 1 and 1,000.";
        qtyValidation.ErrorMessage = "Invalid quantity value.";
        qtyValidation.ShowInput = true;
        qtyValidation.ShowError = true;

        // Save the workbook with the applied validations
        workbook.Save("ValidatedOutput.xlsx");
    }
}