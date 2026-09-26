// Title: Add a rectangle shape that reflects a data‑validation list selection by linking its text to a cell in Aspose.Cells for .NET
// AI Prompts: Insert a rectangle shape at C3, bind its Text property to cell B2 with a formula, and create a dropdown list in B2 using Aspose.Cells for .NET. | Generate a workbook, populate a source range for a list validation, apply the validation to B2, and configure a shape to display the chosen list item. | Create a shape, set its text formula to reference a validated cell, and save the workbook as an .xlsx file with Aspose.Cells.
// Common Searches: Aspose.Cells how to link shape text to a cell value | C# add rectangle shape that displays selected dropdown value in Aspose.Cells | set shape text formula to reference a data validation cell using Aspose.Cells .NET | create data validation list and bind it to a shape in an Excel file with Aspose.Cells | display selected option from a list validation inside a shape in .xlsx using Aspose.Cells
// Tags: add shape linked to cell Aspose.Cells | cell B2 list validation Aspose.Cells | shape text formula reference Aspose.Cells | display dropdown selection in shape .NET | save workbook with linked shape Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example creates a new workbook, fills D1:D5 with option values, adds a list validation to cell B2, inserts a rectangle shape at C3, sets the shape's text to the formula =B2 so it updates with the selected dropdown item, applies basic line styling, and saves the file as ShapeLinkedToValidation.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Access the first worksheet
                var worksheet = workbook.Worksheets[0];

                // Populate a list of values for the validation (e.g., D1:D5)
                worksheet.Cells["D1"].PutValue("Option 1");
                worksheet.Cells["D2"].PutValue("Option 2");
                worksheet.Cells["D3"].PutValue("Option 3");
                worksheet.Cells["D4"].PutValue("Option 4");
                worksheet.Cells["D5"].PutValue("Option 5");

                // Define the cell area for the validation (B2)
                var validationArea = new CellArea
                {
                    StartRow = 1,    // zero‑based row index for B2
                    StartColumn = 1, // zero‑based column index for B2
                    EndRow = 1,
                    EndColumn = 1
                };

                // Add a data‑validation list to cell B2
                int validationIndex = worksheet.Validations.Add(validationArea);
                var validation = worksheet.Validations[validationIndex];
                validation.Type = ValidationType.List;
                validation.InCellDropDown = true;               // Show dropdown arrow
                validation.Formula1 = "D1:D5";                  // Source range for the list

                // Add a rectangle shape to the worksheet
                // Parameters: type, upper left row, upper left column, top, left, height, width
                var shape = worksheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle,
                    2,    // upper left row (zero‑based, row 3)
                    2,    // upper left column (zero‑based, column C)
                    5,    // top offset (pixels)
                    5,    // left offset (pixels)
                    50,   // height (pixels)
                    150); // width (pixels)

                // Set the shape's text to a formula that references cell B2.
                shape.Text = "=B2";

                // Optional styling for better visibility
                shape.Line.Weight = 1.0;
                shape.Line.DashStyle = MsoLineDashStyle.Solid;

                // Save the workbook
                string outputPath = "ShapeLinkedToValidation.xlsx";

                // Ensure the directory exists (if a directory part is present)
                string? outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
