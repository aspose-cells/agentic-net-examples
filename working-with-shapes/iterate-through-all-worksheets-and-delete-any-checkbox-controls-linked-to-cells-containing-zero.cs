// Title: Remove CheckBox Controls Linked to Zero Cells in All Worksheets (Aspose.Cells for .NET)
// Description: Loads a workbook, walks through every worksheet’s CheckBoxCollection, and deletes each CheckBox whose LinkedCell holds the numeric value 0. The backward iteration prevents index errors, and the workbook is saved after cleanup.
// Keywords: Aspose.Cells | C# | CheckBoxCollection | delete checkboxes | linked cell zero | iterate worksheets | Excel form controls | remove shapes | Excel automation | Aspose.Cells example
// Common Searches: Aspose.Cells delete checkboxes linked to zero | remove check box controls from all sheets C# | how to filter check boxes by linked cell value in Excel | iterate worksheets and delete specific shapes Aspose.Cells | C# code to clean up zero‑linked check boxes
// Developer Intent: Programmatically eliminate every CheckBox whose LinkedCell contains the numeric value 0 from all worksheets in an Excel workbook.
// Use Cases: Sanitize a shared template by stripping false‑flag check boxes before distribution. | Reduce visual clutter in financial reports by removing check boxes tied to zero amounts. | Prepare workbooks for publishing or archiving, ensuring no zero‑linked form controls remain.
// AI Prompts: Generate C# code with Aspose.Cells that deletes all CheckBox shapes whose LinkedCell equals 0 across every worksheet. | Show an alternative approach to remove zero‑linked check boxes without using a reverse loop, while keeping collection integrity. | Explain how to log the address and sheet name of each removed CheckBox for audit tracking.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads a workbook, walks through every worksheet’s CheckBoxCollection, and deletes each CheckBox whose LinkedCell holds the numeric value 0. The backward iteration prevents index errors, and the workbook is saved after cleanup.
class DeleteZeroLinkedCheckBoxes
{
    static void Main()
    {
        // Load an existing workbook (replace with actual path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the collection of check boxes on the current worksheet
            CheckBoxCollection checkBoxes = sheet.CheckBoxes;

            // Iterate backwards so that removal does not affect the loop index
            for (int i = checkBoxes.Count - 1; i >= 0; i--)
            {
                CheckBox cb = checkBoxes[i];

                // Only process check boxes that are linked to a cell
                string linkedCell = cb.LinkedCell;
                if (string.IsNullOrEmpty(linkedCell))
                    continue;

                // Retrieve the linked cell
                Cell cell = sheet.Cells[linkedCell];

                // Determine if the cell contains a numeric zero
                if (cell.Value != null && double.TryParse(cell.Value.ToString(), out double numericValue) && numericValue == 0)
                {
                    // Remove the check box from the collection
                    checkBoxes.RemoveAt(i);
                }
            }
        }

        // Save the modified workbook (replace with desired output path)
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
