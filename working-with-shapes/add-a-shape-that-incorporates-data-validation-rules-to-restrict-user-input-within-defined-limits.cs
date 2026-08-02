// Title: Create a ListBox shape with a linked cell and numeric data validation using Aspose.Cells for .NET (C#)
// Description: C# example that builds a workbook, fills A1:A5 with colors, adds a ListBox shape, links its selection to B1, and applies whole‑number validation (1‑10) with custom messages, then saves the file.
// Keywords: Aspose.Cells C# ListBox shape | link shape to cell Aspose.Cells | data validation whole number Aspose.Cells | restrict numeric input in Excel using Aspose | .NET Excel shape validation example | Aspose.Cells add ListBox and validation | Excel form controls with validation C# | Aspose.Cells tutorial data entry
// Common Searches: Aspose.Cells add ListBox shape and link to cell | How to apply numeric validation to a cell linked to a shape in Aspose.Cells | C# example for ListBox input range and data validation with Aspose.Cells | Aspose.Cells restrict user input to a range | Create Excel form controls with validation using Aspose
// Developer Intent: Add a ListBox shape, bind it to a worksheet cell, and enforce whole‑number validation on that cell.
// Use Cases: Build a simple data‑entry form where a color ListBox stores a code in a cell that must be a number between 1 and 10. | Design an interactive worksheet that combines a shape‑based selector with strict numeric validation to prevent invalid entries. | Create multiple ListBox controls, each linked to different cells with distinct validation rules for a guided user experience.
// AI Prompts: Generate C# code with Aspose.Cells to insert a ListBox shape, set its input range, link it to a cell, and apply whole‑number validation from 1 to 10. | Show how to modify the validation to accept dates or custom formulas for a cell linked to a shape in Aspose.Cells. | Provide step‑by‑step instructions for adding several ListBox shapes, each with its own linked cell and unique validation constraints.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// C# example that builds a workbook, fills A1:A5 with colors, adds a ListBox shape, links its selection to B1, and applies whole‑number validation (1‑10) with custom messages, then saves the file.
class ShapeWithValidationExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate a range with list items that will be used by the ListBox shape
        string[] listItems = { "Red", "Green", "Blue", "Yellow", "Purple" };
        for (int i = 0; i < listItems.Length; i++)
        {
            sheet.Cells[i, 0].PutValue(listItems[i]); // Cells A1:A5
        }

        // Add a ListBox shape (acts like a combo box) to the worksheet
        // Parameters: upper left row, upper left column, top, left, width, height
        ListBox listBox = (ListBox)sheet.Shapes.AddListBox(2, 0, 2, 0, 120, 100);

        // Set the range that fills the ListBox with the items defined above
        // false, false indicate A1‑style reference and not locale‑specific
        listBox.SetInputRange("$A$1:$A$5", false, false);

        // Link the selected value of the ListBox to cell B1
        listBox.SetLinkedCell("$B$1", false, false);

        // Add data validation to the linked cell (B1) to restrict numeric input between 1 and 10
        CellArea validationArea = CellArea.CreateCellArea(0, 1, 0, 1); // B1
        int validationIndex = sheet.Validations.Add(validationArea);
        Validation validation = sheet.Validations[validationIndex];

        validation.Type = ValidationType.WholeNumber;
        validation.Operator = OperatorType.Between;
        validation.Formula1 = "1";   // Minimum allowed value
        validation.Formula2 = "10";  // Maximum allowed value

        // Optional user‑friendly messages
        validation.InputTitle = "Enter Number";
        validation.InputMessage = "Please enter a whole number between 1 and 10.";
        validation.ErrorTitle = "Invalid Input";
        validation.ErrorMessage = "The value must be a whole number from 1 to 10.";
        validation.ShowInput = true;
        validation.ShowError = true;
        validation.AlertStyle = ValidationAlertType.Stop;

        // Save the workbook
        workbook.Save("ShapeWithValidation.xlsx");
    }
}
