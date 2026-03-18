using System;
using Aspose.Cells;

namespace AsposeCellsValidationWithCellArea
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the primary validation area (e.g., A1:A10)
            CellArea primaryArea = CellArea.CreateCellArea(0, 0, 9, 0); // rows 0-9, column 0 (A)

            // Add a validation to the collection using the primary area
            int validationIndex = sheet.Validations.Add(primaryArea);
            Validation validation = sheet.Validations[validationIndex];

            // Configure validation properties (list validation as an example)
            validation.Type = ValidationType.List;
            validation.InCellDropDown = true;
            validation.Formula1 = "\"Option1,Option2,Option3\"";

            // Define an additional area (e.g., C1:C5) and add it to the same validation
            CellArea additionalArea = CellArea.CreateCellArea(0, 2, 4, 2); // rows 0-4, column 2 (C)
            validation.AddArea(additionalArea);

            // Optionally add multiple areas in bulk
            CellArea[] bulkAreas = new CellArea[]
            {
                CellArea.CreateCellArea(10, 0, 14, 0), // A11:A15
                CellArea.CreateCellArea(10, 2, 14, 2)  // C11:C15
            };
            // Add bulk areas with intersection and edge checks enabled for safety
            validation.AddAreas(bulkAreas, true, true);

            // Save the workbook (lifecycle rule: save)
            workbook.Save("EnhancedValidationWithCellAreas.xlsx", SaveFormat.Xlsx);
        }
    }
}