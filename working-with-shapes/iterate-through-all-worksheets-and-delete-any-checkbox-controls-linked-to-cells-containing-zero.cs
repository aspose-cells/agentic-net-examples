// Title: Remove CheckBox shapes whose linked cells contain zero across all worksheets with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells remove checkboxes linked to cells with value 0 | iterate all worksheets and delete check box shapes when linked cell equals zero using Aspose.Cells | programmatically delete Excel checkboxes that reference zero cells in .NET | how to clean up zero‑linked checkboxes in an Excel workbook with Aspose.Cells | Aspose.Cells C# delete CheckBox if linked cell is numeric zero
// Tags: Aspose.Cells check box collection manipulation | delete Excel check boxes based on linked cell content | iterate worksheets to clean up shape objects | zero-value linked cell detection in Aspose.Cells | C# programmatic removal of Excel shapes

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, loops through each worksheet, examines every CheckBox's LinkedCell, determines if the linked cell's value is numeric zero (including string parsing), removes the CheckBox when the condition is met, and saves the updated workbook.
class Program
{
    static void Main(string[] args)
    {
        // Validate input arguments
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: Program.exe <input.xlsx> <output.xlsx>");
            return;
        }

        string inputPath = args[0];
        string outputPath = args[1];

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Console.WriteLine($"Loaded workbook: {inputPath}");

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Console.WriteLine($"Processing worksheet: {sheet.Name}");

                // Get the collection of check boxes on the current sheet
                CheckBoxCollection checkBoxes = sheet.CheckBoxes;

                // Iterate backwards so that removal does not affect the loop index
                for (int i = checkBoxes.Count - 1; i >= 0; i--)
                {
                    CheckBox cb = checkBoxes[i];
                    string linkedCellAddress = cb.LinkedCell;

                    // If the checkbox is not linked to any cell, skip it
                    if (string.IsNullOrWhiteSpace(linkedCellAddress))
                        continue;

                    // Retrieve the cell that the checkbox is linked to
                    Cell linkedCell = sheet.Cells[linkedCellAddress];
                    if (linkedCell == null || linkedCell.Value == null)
                        continue;

                    // Determine if the cell's value is numeric zero
                    bool isZero = false;
                    object cellValue = linkedCell.Value;

                    if (cellValue is double d)
                    {
                        isZero = d == 0.0;
                    }
                    else if (cellValue is int iVal)
                    {
                        isZero = iVal == 0;
                    }
                    else if (cellValue is long lVal)
                    {
                        isZero = lVal == 0L;
                    }
                    else
                    {
                        // Attempt to parse any other representation as a double
                        if (double.TryParse(cellValue.ToString(), out double parsed))
                        {
                            isZero = parsed == 0.0;
                        }
                    }

                    // If the linked cell contains zero, delete the checkbox
                    if (isZero)
                    {
                        Console.WriteLine($"Deleting CheckBox linked to {linkedCellAddress} (value=0) on sheet \"{sheet.Name}\".");
                        checkBoxes.RemoveAt(i);
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved after removal: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred while processing the workbook:");
            Console.WriteLine(ex.Message);
        }
    }
}
