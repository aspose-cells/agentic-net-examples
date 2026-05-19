using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsCheckboxCleanup
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load the workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the collection of check boxes on the current worksheet
                CheckBoxCollection checkBoxes = sheet.CheckBoxes;

                // Iterate backwards so that removal does not affect the index order
                for (int i = checkBoxes.Count - 1; i >= 0; i--)
                {
                    CheckBox cb = checkBoxes[i];

                    // Process only check boxes that are linked to a cell
                    if (!string.IsNullOrEmpty(cb.LinkedCell))
                    {
                        // Retrieve the linked cell
                        Cell linkedCell = sheet.Cells[cb.LinkedCell];

                        // Ensure the cell has a numeric value equal to zero
                        if (linkedCell != null && linkedCell.Value != null && 
                            double.TryParse(linkedCell.Value.ToString(), out double numericValue) &&
                            numericValue == 0)
                        {
                            // Remove the check box from the collection
                            checkBoxes.RemoveAt(i);
                        }
                    }
                }
            }

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}