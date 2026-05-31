using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeLinkedCellValidation
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Example: add a few shapes with linked cells
            Shape shape1 = sheet.Shapes.AddRectangle(1, 1, 100, 100, 0, 0);
            shape1.LinkedCell = "$A$2"; // valid link

            Shape shape2 = sheet.Shapes.AddRectangle(2, 2, 100, 100, 0, 0);
            shape2.LinkedCell = ""; // invalid (empty)

            Shape shape3 = sheet.Shapes.AddRectangle(3, 3, 100, 100, 0, 0);
            shape3.LinkedCell = "$B$5"; // will be validated for non‑empty value

            // Populate some data for demonstration
            sheet.Cells["A2"].PutValue("Hello"); // non‑empty
            sheet.Cells["B5"].PutValue("");      // empty

            // Iterate through all shapes in the worksheet
            foreach (Shape shape in sheet.Shapes)
            {
                // Retrieve the linked cell address (property rule)
                string linkedAddress = shape.LinkedCell;

                // Check if the linked address is non‑empty and well‑formed
                if (string.IsNullOrWhiteSpace(linkedAddress))
                {
                    Console.WriteLine($"Shape '{shape.Name}' has an empty LinkedCell reference.");
                    continue; // skip adding validation for this shape
                }

                // Obtain the cell object using the address (A1 notation)
                Cell linkedCell;
                try
                {
                    linkedCell = sheet.Cells[linkedAddress];
                }
                catch (Exception)
                {
                    Console.WriteLine($"Shape '{shape.Name}' has an invalid LinkedCell address: {linkedAddress}");
                    continue;
                }

                // Create a custom validation that ensures the linked cell is not empty
                // Validation rule: LEN(TRIM(address))>0
                string formula = $"=LEN(TRIM({linkedAddress}))>0";

                // Define the area for the validation (single cell)
                CellArea area = CellArea.CreateCellArea(linkedCell.Row, linkedCell.Column, linkedCell.Row, linkedCell.Column);
                int validationIndex = sheet.Validations.Add(area);
                Validation validation = sheet.Validations[validationIndex];

                validation.Type = ValidationType.Custom;
                validation.Formula1 = formula;
                validation.ShowError = true;
                validation.ErrorTitle = "Invalid Linked Cell";
                validation.ErrorMessage = $"The linked cell {linkedAddress} must contain a non‑empty value.";
                validation.AlertStyle = ValidationAlertType.Stop;

                Console.WriteLine($"Added validation to linked cell {linkedAddress} for shape '{shape.Name}'.");
            }

            // Save the workbook (save rule)
            workbook.Save("ShapeLinkedCellValidation.xlsx", SaveFormat.Xlsx);
        }
    }
}