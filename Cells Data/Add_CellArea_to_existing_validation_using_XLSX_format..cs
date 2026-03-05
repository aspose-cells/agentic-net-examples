using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        // Define the initial validation area (cell A1)
        CellArea initialArea = CellArea.CreateCellArea(0, 0, 0, 0);

        // Add a validation to the collection for the initial area
        int validationIndex = ws.Validations.Add(initialArea);
        Validation validation = ws.Validations[validationIndex];

        // Configure the validation (list type with three options)
        validation.Type = ValidationType.List;
        validation.Formula1 = "\"Option1,Option2,Option3\"";

        // Define an additional area (cells B2:C3) to apply the same validation
        CellArea additionalArea = CellArea.CreateCellArea(1, 1, 2, 2);

        // Add the additional area to the existing validation
        validation.AddArea(additionalArea);

        // Save the workbook in XLSX format
        wb.Save("ValidationWithMultipleAreas.xlsx", SaveFormat.Xlsx);
    }
}