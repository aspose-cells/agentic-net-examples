// Title: C# – Add a CheckBox shape linked to a volatile NOW cell and auto‑update it with Aspose.Cells
// Description: Creates a new workbook, inserts a CheckBox shape on the first worksheet, links the shape to cell B2, assigns the volatile NOW formula to that cell, runs an initial calculation, pauses, recalculates to get a fresh timestamp, updates the CheckBox state, and saves the file as ShapeLinkedCellVolatile.xlsx.
// Keywords: Aspose.Cells C# | add checkbox shape | link shape to cell | volatile NOW formula | recalculate workbook | update shape selected value | Excel automation | dynamic timestamp
// Common Searches: Aspose.Cells link checkbox to cell | how to recalculate NOW formula in Aspose.Cells | auto‑update shape after volatile function | C# example for checkbox linked cell | refresh Excel shape after delay
// Developer Intent: Insert a CheckBox, bind it to a cell containing a volatile NOW formula, trigger recalculation, and synchronize the shape’s state automatically.
// Use Cases: Display a live timestamp in a report where a checkbox reflects the most recent calculation. | Validate that shape selections stay in sync with volatile data during automated workbook processing. | Create audit worksheets that automatically toggle a checkbox based on the latest formula result.
// AI Prompts: Generate C# code using Aspose.Cells to add a checkbox, link it to cell C5, set a TODAY formula, wait, recalculate, and sync the checkbox state. | Explain the workflow of volatile functions in Aspose.Cells and how to force shape updates after Workbook.CalculateFormula. | Provide a unit‑test outline that verifies automatic updates of a shape linked to a volatile function in Aspose.Cells for .NET.

using System;
using System.IO;
using System.Threading;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, inserts a CheckBox shape on the first worksheet, links the shape to cell B2, assigns the volatile NOW formula to that cell, runs an initial calculation, pauses, recalculates to get a fresh timestamp, updates the CheckBox state, and saves the file as ShapeLinkedCellVolatile.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a CheckBox shape to the worksheet
            // Parameters: upper left row, upper left column, row offset, column offset, width, height
            CheckBox checkBox = (CheckBox)worksheet.Shapes.AddCheckBox(1, 1, 0, 0, 100, 30);

            // Link the CheckBox to cell B2 (A1 style, locale aware)
            checkBox.SetLinkedCell("$B$2", false, true);

            // Set a volatile function (NOW) in the linked cell
            worksheet.Cells["B2"].SetFormula("=NOW()", null);

            // Perform initial calculation
            workbook.CalculateFormula();

            Console.WriteLine("Initial linked cell value: " + worksheet.Cells["B2"].Value);

            // Simulate a delay to allow the volatile function to change
            Thread.Sleep(2000);

            // Recalculate to update the volatile function result
            workbook.CalculateFormula();

            Console.WriteLine("After recalculation linked cell value: " + worksheet.Cells["B2"].Value);

            // Update the CheckBox's selected state based on the linked cell
            checkBox.UpdateSelectedValue();

            // Save the workbook
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ShapeLinkedCellVolatile.xlsx");
            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved to: " + outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
