// Title: Create a ListBox Shape with Linked Cell and Whole‑Number Validation in Aspose.Cells for .NET (C#)
// Description: This C# example shows how to add a ListBox shape to a worksheet, populate it from range A1:A5, link the selected item to cell B2, and enforce a whole‑number data‑validation rule (1‑100) on that linked cell before saving the workbook as ShapeWithValidation.xlsx.
// Keywords: Aspose.Cells ListBox shape C# | shape linked cell data validation | whole number validation Aspose.Cells | Excel form control with validation .NET | add ListBox shape programmatically | restrict user input Aspose.Cells | C# Excel shape validation example | Aspose.Cells shape UI control
// Common Searches: how to add a ListBox shape in Aspose.Cells | link ListBox shape to a cell using C# | apply numeric data validation to a linked cell Aspose.Cells | restrict ListBox selection to numbers 1 to 100 | Aspose.Cells shape with validation example
// Developer Intent: Insert a ListBox shape, bind its value to a worksheet cell, and apply a numeric range validation to that cell.
// Use Cases: Design interactive spreadsheet forms where users pick an item from a drop‑down list and the result must be a valid numeric code. | Combine visual controls with data‑validation rules to prevent out‑of‑range entries in financial or inventory templates. | Create a reusable UI component that writes a validated number back to the sheet for downstream calculations or macros.
// AI Prompts: Generate C# code with Aspose.Cells to add a ComboBox shape linked to cell C5 and enforce a date validation between 01/01/2024 and 12/31/2024. | Explain how to replace the static A1:A5 source range with a dynamic named range while keeping the whole‑number validation on the linked cell. | Provide steps to programmatically read the value selected in the ListBox after the workbook is opened in Excel and raise a custom warning if it falls outside the allowed range.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsValidationWithShape
{
    // This C# example shows how to add a ListBox shape to a worksheet, populate it from range A1:A5, link the selected item to cell B2, and enforce a whole‑number data‑validation rule (1‑100) on that linked cell before saving the workbook as ShapeWithValidation.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Prepare data that will be used for the shape list
            // -------------------------------------------------
            // Fill cells A1:A5 with sample list values
            for (int i = 0; i < 5; i++)
            {
                worksheet.Cells[i, 0].PutValue($"Item {i + 1}");
            }

            // -------------------------------------------------
            // Add a ListBox shape to the worksheet
            // -------------------------------------------------
            // Parameters: upperLeftRow, upperLeftColumn, top, left, width, height
            Shape listBox = worksheet.Shapes.AddListBox(2, 0, 2, 0, 130, 130);

            // Set the range that supplies the list items for the ListBox
            // Using A1:A5 as the source range
            listBox.SetInputRange("$A$1:$A$5", false, false);

            // Link the selected value of the ListBox to cell B2
            listBox.SetLinkedCell("$B$2", false, false);

            // -------------------------------------------------
            // Add data validation to the linked cell (B2)
            // -------------------------------------------------
            // Define the cell area for validation (B2)
            CellArea validationArea = CellArea.CreateCellArea(1, 1, 1, 1); // Row 1, Column 1 (zero‑based)

            // Add a new validation to the worksheet's validation collection
            int validationIndex = worksheet.Validations.Add(validationArea);
            Validation validation = worksheet.Validations[validationIndex];

            // Configure the validation: whole number between 1 and 100
            validation.Type = ValidationType.WholeNumber;
            validation.Operator = OperatorType.Between;
            validation.Formula1 = "1";
            validation.Formula2 = "100";

            // Optional UI messages
            validation.ShowInput = true;
            validation.InputTitle = "Enter Number";
            validation.InputMessage = "Please enter a whole number between 1 and 100.";
            validation.ShowError = true;
            validation.ErrorTitle = "Invalid Input";
            validation.ErrorMessage = "The value must be a whole number between 1 and 100.";
            validation.AlertStyle = ValidationAlertType.Stop;

            // -------------------------------------------------
            // Save the workbook
            // -------------------------------------------------
            workbook.Save("ShapeWithValidation.xlsx");
        }
    }
}
