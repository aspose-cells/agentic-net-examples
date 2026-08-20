// Title: Link a Shape to an External Workbook Cell and Verify the Value with Aspose.Cells for .NET
// Description: Demonstrates how to create an external workbook, add a rectangle shape to a main workbook, link the shape to cell B2, set a formula that references the external file, register the link via the ExternalLinks collection, refresh data with UpdateLinkedDataSource, recalculate formulas, and confirm that the shape displays the external value using Aspose.Cells for C#.
// Keywords: Aspose.Cells shape link external workbook | C# rectangle shape linked cell | ExternalLinks collection Aspose.Cells | UpdateLinkedDataSource method | verify shape displays external value | link shape to cell formula | cross‑workbook reference Aspose.Cells
// Common Searches: Aspose.Cells link shape to external workbook cell | C# shape linked cell formula external file | How to update external links in Aspose.Cells | Verify shape value after external formula calculation | Add rectangle shape and link to cell using Aspose.Cells
// Developer Intent: Create a shape whose linked cell pulls data from another workbook, refresh the link, and ensure the shape shows the updated value.
// Use Cases: Build dashboards where shapes reflect live data from a separate source workbook. | Generate reports that automatically update shape captions when the linked data workbook changes. | Automate testing of external links by reading the linked cell after formula recalculation.
// AI Prompts: Write C# code with Aspose.Cells to add a rectangle shape, link it to cell B2, set B2's formula to reference an external workbook, register the external link, refresh data, recalculate formulas, and output the linked cell value. | Explain the role of the ExternalLinks collection and the UpdateLinkedDataSource method in refreshing shape‑linked cells from an external file using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExternalShapeLinkDemo
{
    // Demonstrates how to create an external workbook, add a rectangle shape to a main workbook, link the shape to cell B2, set a formula that references the external file, register the link via the ExternalLinks collection, refresh data with UpdateLinkedDataSource, recalculate formulas, and confirm that the shape displays the external value using Aspose.Cells for C#.
    class Program
    {
        static void Main()
        {
            // Wrap the whole process in a try-catch to handle unexpected errors gracefully
            try
            {
                // ---------- Create external workbook ----------
                Workbook externalWb = new Workbook();
                Worksheet externalWs = externalWb.Worksheets[0];
                externalWs.Name = "Sheet1";

                // Put a test value in A1 of the external workbook
                externalWs.Cells["A1"].PutValue("External Value");

                // Save external workbook (required for external link to resolve)
                string externalFile = "ExternalData.xlsx";
                externalWb.Save(externalFile);

                // ---------- Create main workbook ----------
                Workbook mainWb = new Workbook();
                Worksheet mainWs = mainWb.Worksheets[0];
                mainWs.Name = "MainSheet";

                // Add a rectangle shape to the main worksheet
                // Parameters: upper left row, upper left column, top offset, left offset, height, width
                Shape shape = mainWs.Shapes.AddRectangle(2, 2, 0, 0, 100, 200);

                // Link the shape to cell B2 (no $ signs – LinkedCell expects a simple address)
                shape.LinkedCell = "B2";

                // Set formula in the linked cell to reference the external workbook
                mainWs.Cells["B2"].Formula = $"='[{externalFile}]Sheet1'!A1";

                // Register the external link in the workbook's external links collection
                int linkIndex = mainWb.Worksheets.ExternalLinks.Add(externalFile, new string[] { "Sheet1" });
                ExternalLink extLink = mainWb.Worksheets.ExternalLinks[linkIndex];
                Console.WriteLine($"Added external link with DataSource: {extLink.DataSource}");

                // Update linked data source so that the main workbook fetches the latest value
                mainWb.UpdateLinkedDataSource(new Workbook[] { externalWb });

                // Recalculate formulas to reflect the external value
                mainWb.CalculateFormula();

                // Verify that the shape's linked cell now contains the external value
                string linkedCellAddress = shape.LinkedCell;
                string linkedCellValue = mainWs.Cells[linkedCellAddress].StringValue;
                Console.WriteLine($"Shape linked cell ({linkedCellAddress}) value: {linkedCellValue}");

                // Save the main workbook
                string mainFile = "MainWorkbookWithShapeLink.xlsx";
                mainWb.Save(mainFile);
                Console.WriteLine($"Main workbook saved as {mainFile}");

                // Clean up
                externalWb.Dispose();
                mainWb.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
