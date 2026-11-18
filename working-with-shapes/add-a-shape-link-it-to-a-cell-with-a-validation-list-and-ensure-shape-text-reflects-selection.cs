// Title: Create a rectangle shape linked to a data‑validation list cell and display the selected value with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add rectangle shape that shows the value of a data validation list cell | How to bind shape text to a dropdown cell using Aspose.Cells for .NET | Set shape formula to a validated cell in Excel with Aspose.Cells C# example | Link Excel shape to cell B2 that has a data validation list using Aspose.Cells | Create data validation list and connect shape text to it in C# Aspose.Cells
// Tags: rectangle shape data validation Aspose.Cells C# | shape text formula linked to cell Aspose.Cells | bind shape to validated cell value C# | add shape with formula =B2 using Aspose.Cells | Excel shape linked to dropdown selection Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace Example
{
    // The example creates a workbook, fills D1:D3 with list items, adds a data‑validation list to cell B2, inserts a rectangle shape, sets the shape's Text property to the formula "=B2" so it reflects the selected dropdown value, and saves the file as an XLSX workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------------------------------------------------------
                // 1. Prepare the list source for data validation (e.g., D1:D3)
                sheet.Cells["D1"].PutValue("Option 1");
                sheet.Cells["D2"].PutValue("Option 2");
                sheet.Cells["D3"].PutValue("Option 3");

                // -------------------------------------------------------------------
                // 2. Add a data validation list to cell B2 that points to the list above
                // Define the cell area for B2 (row 1, column 1 in zero‑based indexing)
                CellArea area = new CellArea
                {
                    StartRow = 1,
                    EndRow = 1,
                    StartColumn = 1,
                    EndColumn = 1
                };
                int validationIndex = sheet.Validations.Add(area);
                Validation validation = sheet.Validations[validationIndex];
                validation.Type = ValidationType.List;
                validation.Operator = OperatorType.Equal;
                validation.Formula1 = "=$D$1:$D$3"; // reference the list range
                validation.ShowError = true;
                validation.ErrorTitle = "Invalid Selection";
                validation.ErrorMessage = "Please select a value from the list.";
                validation.InCellDropDown = true;

                // -------------------------------------------------------------------
                // 3. Insert a shape (rectangle) onto the worksheet
                // AddShape returns the created Shape object directly
                Shape shape = sheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle, // shape type
                    5,    // upper left row
                    0,    // upper left column
                    0,    // upper left row offset (pixels)
                    0,    // upper left column offset (pixels)
                    200,  // shape width (pixels)
                    50    // shape height (pixels)
                );

                // -------------------------------------------------------------------
                // 4. Link the shape's displayed text to the value of the validated cell (B2)
                shape.Text = "=B2";

                // -------------------------------------------------------------------
                // 5. Save the workbook to a file
                string outputPath = "ShapeLinkedToValidation.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
