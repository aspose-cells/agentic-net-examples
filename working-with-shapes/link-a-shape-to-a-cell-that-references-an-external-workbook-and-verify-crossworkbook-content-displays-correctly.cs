// Title: Aspose.Cells for .NET: Link a Shape to an External Workbook Cell and Verify the Value (C#)
// Description: C# example that creates an external workbook, adds a rectangle shape to a main workbook, links the shape to cell B2, sets B2’s formula to reference the external workbook, registers the external link, refreshes the data source, calculates formulas, and confirms the shape displays the external value.
// Keywords: Aspose.Cells | C# | .NET | shape linked cell | external workbook reference | cross‑workbook formula | UpdateLinkedDataSource | ExternalLinks collection | rectangle shape linking | linked shape verification | spreadsheet automation
// Common Searches: link shape to cell from another workbook Aspose.Cells | Aspose.Cells external link for shape LinkedCell | update linked data source for cross‑workbook formulas .NET | verify shape displays external workbook value | C# Aspose.Cells external workbook example
// Developer Intent: The developer needs to connect a shape to a cell that pulls data from an external workbook and confirm that the shape shows the retrieved value.
// Use Cases: Generate a source workbook, write a value to A1, and save it as a file. | Create a destination workbook, add a rectangle shape, set its LinkedCell to B2, and assign B2 a formula that points to the source workbook’s A1. | Add the source file to the destination workbook’s ExternalLinks collection, call UpdateLinkedDataSource, and run CalculateFormula. | Read the value of B2 (or the shape’s LinkedCell) to ensure the external data appears correctly. | Optionally save the destination workbook to preserve the linked shape configuration.
// AI Prompts: Write C# code with Aspose.Cells that links a rectangle shape to a cell referencing an external workbook and updates the linked data source. | Explain step‑by‑step how to verify that a shape reflects a value from another workbook after formula calculation in Aspose.Cells. | Provide troubleshooting advice when a linked shape shows a #REF! or does not update after changing the external workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExternalShapeLinkDemo
{
    // C# example that creates an external workbook, adds a rectangle shape to a main workbook, links the shape to cell B2, sets B2’s formula to reference the external workbook, registers the external link, refreshes the data source, calculates formulas, and confirms the shape displays the external value.
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create the external workbook that will serve as data source.
            // -----------------------------------------------------------------
            Workbook externalWb = new Workbook();
            Worksheet externalWs = externalWb.Worksheets[0];
            externalWs.Name = "Sheet1";
            externalWs.Cells["A1"].PutValue("Cross‑Workbook Value");
            // Save the external workbook to disk (required for linking).
            string externalFileName = "ExternalData.xlsx";
            externalWb.Save(externalFileName);

            // -----------------------------------------------------------------
            // 2. Create the main workbook where the shape will be placed.
            // -----------------------------------------------------------------
            Workbook mainWb = new Workbook();
            Worksheet mainWs = mainWb.Worksheets[0];
            mainWs.Name = "MainSheet";

            // -----------------------------------------------------------------
            // 3. Add a rectangle shape and link it to a cell (B2) in the main sheet.
            // -----------------------------------------------------------------
            Shape rect = mainWs.Shapes.AddRectangle(1, 1, 100, 100, 0, 0);
            rect.Name = "LinkedRectangle";
            rect.LinkedCell = "$B$2"; // The shape will display the value of B2.

            // -----------------------------------------------------------------
            // 4. Define a formula in B2 that references the external workbook.
            // -----------------------------------------------------------------
            mainWs.Cells["B2"].Formula = $"=[{externalFileName}]Sheet1!A1";

            // -----------------------------------------------------------------
            // 5. Register the external link in the workbook's ExternalLinks collection.
            // -----------------------------------------------------------------
            // This ensures the link is recognized by Aspose.Cells.
            mainWb.Worksheets.ExternalLinks.Add(externalFileName, new string[] { "Sheet1" });

            // -----------------------------------------------------------------
            // 6. Update the linked data source so the formula can retrieve the latest value.
            // -----------------------------------------------------------------
            // Load the external workbook again (simulating an external source that may have changed).
            Workbook externalWbForUpdate = new Workbook(externalFileName);
            mainWb.UpdateLinkedDataSource(new Workbook[] { externalWbForUpdate });

            // -----------------------------------------------------------------
            // 7. Calculate formulas to evaluate the external reference.
            // -----------------------------------------------------------------
            mainWb.CalculateFormula();

            // -----------------------------------------------------------------
            // 8. Verify that the shape reflects the external value via its linked cell.
            // -----------------------------------------------------------------
            string linkedCellValue = mainWs.Cells["B2"].StringValue;
            Console.WriteLine($"Value in linked cell B2 (should come from external workbook): {linkedCellValue}");
            Console.WriteLine($"Shape '{rect.Name}' is linked to cell: {rect.LinkedCell}");

            // -----------------------------------------------------------------
            // 9. Save the main workbook (optional, demonstrates full lifecycle).
            // -----------------------------------------------------------------
            mainWb.Save("MainWorkbook_WithLinkedShape.xlsx");
        }
    }
}
