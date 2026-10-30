// Title: Create a rectangle shape linked to a data‑validation list cell and display the selected option with Aspose.Cells for .NET (C#)
// Common Searches: how to bind a shape's text to a dropdown cell using Aspose.Cells C# | Aspose.Cells C# add rectangle shape that shows the selected value from a data validation list | set shape text to a formula like =B2 with Aspose.Cells .NET | display data validation list selection in a shape in an Excel workbook using Aspose.Cells | add data validation list to a cell and link a shape to it in C# Aspose.Cells
// Tags: add rectangle shape Aspose.Cells | link shape text to cell formula Aspose.Cells | create data validation list Aspose.Cells | display dropdown selection in shape Aspose.Cells | C# Aspose.Cells shape text binding

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example creates a new workbook, adds a list‑type data validation to cell B2 with three options, inserts a rectangle shape at a specified location, sets the shape's text to the formula "=B2" so it automatically reflects the selected dropdown value, and saves the file as ShapeLinkedToValidation.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Access the first worksheet
                var worksheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // 1. Add a data‑validation list to cell B2
                // -------------------------------------------------
                // Define the cell area for B2 (row 1, column 1 – zero‑based)
                var cellArea = new CellArea
                {
                    StartRow = 1,
                    StartColumn = 1,
                    EndRow = 1,
                    EndColumn = 1
                };

                // Add validation for the defined area
                int validationIndex = worksheet.Validations.Add(cellArea);
                var validation = worksheet.Validations[validationIndex];
                validation.Type = ValidationType.List;
                validation.Operator = OperatorType.None;

                // List items as a comma‑separated string enclosed in quotes
                validation.Formula1 = "\"Option1,Option2,Option3\"";

                // -------------------------------------------------
                // 2. Insert a shape (rectangle) onto the sheet
                // -------------------------------------------------
                int upperLeftRow = 4;      // zero‑based row index where the shape starts
                int upperLeftColumn = 1;   // zero‑based column index where the shape starts
                int top = 0;               // offset in pixels from the top of the cell
                int left = 0;              // offset in pixels from the left of the cell
                int height = 60;           // shape height in pixels
                int width = 200;           // shape width in pixels

                var shape = worksheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle,
                    upperLeftRow,
                    upperLeftColumn,
                    top,
                    left,
                    height,
                    width);

                // Optional: give the shape a meaningful name
                shape.Name = "SelectionDisplay";

                // -------------------------------------------------
                // 3. Link the shape's text to the validated cell (B2)
                // -------------------------------------------------
                // Setting the Text to a formula makes the shape display the cell's value.
                shape.Text = "=B2";

                // -------------------------------------------------
                // 4. Save the workbook
                // -------------------------------------------------
                string outputPath = "ShapeLinkedToValidation.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
