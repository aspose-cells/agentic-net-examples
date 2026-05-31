using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data (header + rows)
        worksheet.Cells["A1"].PutValue("Quantity");
        worksheet.Cells["B1"].PutValue("Item");
        worksheet.Cells["A2"].PutValue(10);
        worksheet.Cells["B2"].PutValue("Apple");
        worksheet.Cells["A3"].PutValue(55);
        worksheet.Cells["B3"].PutValue("Banana");

        // Add a data validation rule: whole numbers between 1 and 100 for the Quantity column
        Validation validation = worksheet.Validations[worksheet.Validations.Add()];
        validation.Type = ValidationType.WholeNumber;
        validation.Operator = OperatorType.Between;
        validation.Formula1 = "1";
        validation.Formula2 = "100";
        validation.InputMessage = "Enter a quantity between 1 and 100.";
        validation.ErrorMessage = "Invalid quantity.";
        validation.ShowInput = true;
        validation.ShowError = true;

        // Apply the validation to cells A2:A10
        CellArea area = new CellArea
        {
            StartRow = 1,   // Row index 1 = A2
            StartColumn = 0,
            EndRow = 9,     // Row index 9 = A10
            EndColumn = 0
        };
        validation.AddArea(area);

        // Configure JSON save options
        JsonSaveOptions saveOptions = new JsonSaveOptions
        {
            ExportNestedStructure = false, // flat structure
            HasHeaderRow = true,           // first row contains headers
            SkipEmptyRows = true,          // omit empty rows
            ExportEmptyCells = false       // do not include null for empty cells
        };

        // Save the workbook as a JSON file using the configured options
        string outputPath = "output.json";
        workbook.Save(outputPath, saveOptions);

        Console.WriteLine($"Workbook exported to JSON file: {outputPath}");
    }
}