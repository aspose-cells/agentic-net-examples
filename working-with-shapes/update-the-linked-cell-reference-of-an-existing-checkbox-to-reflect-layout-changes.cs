// Title: Change a CheckBox LinkedCell in an Excel workbook using Aspose.Cells for .NET (C#)
// Description: Loads a workbook, finds the first CheckBox form control on the first worksheet, updates its LinkedCell (or uses SetLinkedCell) to a new address such as $B$2, and saves the file. Includes a guard for worksheets without CheckBoxes.
// Keywords: Aspose.Cells | C# | CheckBox | LinkedCell | update linked cell | Excel form control | SetLinkedCell method | modify shape reference | Aspose.Cells for .NET | programmatic Excel | cell address change | layout adjustment
// Common Searches: Aspose.Cells change CheckBox linked cell | C# set LinkedCell property for Excel CheckBox | SetLinkedCell Aspose.Cells example | How to modify Excel CheckBox reference programmatically | Update form control linked cell after inserting rows
// Developer Intent: Assign a new cell address to an existing CheckBox control in a worksheet.
// Use Cases: After repositioning form controls in a generated report, update each CheckBox's LinkedCell to point to the correct data cell. | When rows or columns are inserted programmatically, adjust CheckBox linked cells to maintain accurate data binding. | Batch‑process multiple CheckBoxes to map them to a predefined set of data cells based on their layout.
// AI Prompts: Generate C# code with Aspose.Cells that iterates over all CheckBoxes on a worksheet and sets each LinkedCell to the cell in column B of the same row. | Show how to call SetLinkedCell with appropriate flags to change a CheckBox's linked cell while preserving its existing formatting and properties. | Explain a safe pattern for updating a CheckBox's LinkedCell when the worksheet may contain zero CheckBoxes, avoiding null‑reference exceptions.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads a workbook, finds the first CheckBox form control on the first worksheet, updates its LinkedCell (or uses SetLinkedCell) to a new address such as $B$2, and saves the file. Includes a guard for worksheets without CheckBoxes.
class UpdateCheckBoxLinkedCell
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Verify that the worksheet contains at least one CheckBox
        if (sheet.CheckBoxes.Count > 0)
        {
            // Retrieve the first CheckBox in the collection
            CheckBox checkBox = sheet.CheckBoxes[0];

            // Update the linked cell reference to reflect the new layout
            // Example: link the CheckBox to cell B2
            checkBox.LinkedCell = "$B$2";

            // If you prefer to use the SetLinkedCell method with explicit flags:
            // checkBox.SetLinkedCell("$B$2", false, true);
        }

        // Save the workbook with the updated linked cell reference
        workbook.Save("output.xlsx");
    }
}
