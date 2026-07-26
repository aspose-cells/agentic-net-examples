// Title: C# – Link a Label Shape to a VLOOKUP Result Cell and Keep It Updated with Aspose.Cells
// Description: Demonstrates how to create a workbook, build a lookup table, insert a VLOOKUP formula, add a label shape, bind the shape to the formula cell via the LinkedCell property, force the shape to read the current value with UpdateSelectedValue, and save the file. The shape automatically reflects any changes to the VLOOKUP result.
// Keywords: Aspose.Cells C# | label shape LinkedCell | VLOOKUP shape binding | update shape after recalculation | Excel automation dynamic text | bind shape to cell value | Excel dashboard label | cell‑shape synchronization | Aspose.Cells example
// Common Searches: Aspose.Cells link shape to cell C# | label shape display VLOOKUP result | update shape text after formula change | bind Excel shape to formula cell | Aspose.Cells LinkedCell property usage
// Developer Intent: The developer wants to attach a label shape to a cell that contains a VLOOKUP formula so the shape shows the lookup result and stays synchronized when the workbook recalculates.
// Use Cases: Design an Excel dashboard where a label shape always shows the latest lookup value as source data changes. | Generate reports that use shapes as visual markers linked to calculated cells, ensuring the markers reflect current formula outcomes. | Create interactive workbooks where shapes act as dynamic captions tied to VLOOKUP results for clearer end‑user presentation.
// AI Prompts: Write C# code using Aspose.Cells to add a label shape, link it to a cell with a VLOOKUP formula, and display the current result. | Show how to keep a shape linked to a formula cell updated automatically after workbook recalculation with Aspose.Cells for .NET. | Provide an example that links multiple label shapes to different VLOOKUP result cells and refreshes their displayed values in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeLinkedCellDemo
{
    // Demonstrates how to create a workbook, build a lookup table, insert a VLOOKUP formula, add a label shape, bind the shape to the formula cell via the LinkedCell property, force the shape to read the current value with UpdateSelectedValue, and save the file. The shape automatically reflects any changes to the VLOOKUP result.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // 1. Prepare data for VLOOKUP
            // -------------------------------------------------
            // Table range A1:B5 (lookup key in column A, value in column B)
            sheet.Cells["A1"].Value = "Key";
            sheet.Cells["B1"].Value = "Value";
            sheet.Cells["A2"].Value = "Apple";
            sheet.Cells["B2"].Value = 10;
            sheet.Cells["A3"].Value = "Banana";
            sheet.Cells["B3"].Value = 20;
            sheet.Cells["A4"].Value = "Cherry";
            sheet.Cells["B4"].Value = 30;
            sheet.Cells["A5"].Value = "Date";
            sheet.Cells["B5"].Value = 40;

            // -------------------------------------------------
            // 2. Insert VLOOKUP formula in cell D2
            //    =VLOOKUP("Banana", $A$2:$B$5, 2, FALSE)
            // -------------------------------------------------
            sheet.Cells["D2"].Formula = "=VLOOKUP(\"Banana\", $A$2:$B$5, 2, FALSE)";

            // -------------------------------------------------
            // 3. Add a label shape that will display the result of the VLOOKUP
            // -------------------------------------------------
            // Parameters: upper left row, upper left column, height, width, upper left row offset, upper left column offset
            Label label = (Label)sheet.Shapes.AddLabel(2, 3, 100, 30, 0, 0);
            // Link the label to the cell containing the VLOOKUP result
            label.LinkedCell = "$D$2";

            // -------------------------------------------------
            // 4. Force the shape to read the linked cell value (optional but ensures up‑to‑date display)
            // -------------------------------------------------
            label.UpdateSelectedValue();

            // -------------------------------------------------
            // 5. Save the workbook
            // -------------------------------------------------
            workbook.Save("ShapeLinkedToVlookup.xlsx");
        }
    }
}
