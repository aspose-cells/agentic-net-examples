// Title: Link a rectangle shape to a cell that references an external workbook and verify the displayed value using Aspose.Cells for .NET
// AI Prompts: Create an external workbook named External.xlsx, write a value to cell A1, and save it. | In a new workbook, add a formula that references '[External.xlsx]Sheet1'!A1, insert a rectangle shape, and set the shape’s internal hyperlink to the formula cell (e.g., B2). | Force formula calculation, read the value from the linked cell, and output it to confirm the external reference is resolved.
// Common Searches: Aspose.Cells C# link shape to a cell that contains an external workbook formula | how to set an internal hyperlink on a shape that points to a cell referencing another workbook | verify that a shape hyperlink displays the value from an external workbook in .NET | C# Aspose.Cells create rectangle shape and assign hyperlink to cell B2 | recalculate formulas that include external workbook references using Aspose.Cells
// Tags: shape creation Aspose.Cells | internal hyperlink from shape to cell | external workbook formula reference | calculate workbook with external links | verify external cell value Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates creating an external workbook, referencing its A1 cell from a main workbook via a formula, adding a rectangle shape linked to that cell through an internal hyperlink, recalculating formulas, and saving both workbooks while confirming the external value is correctly displayed.
class ShapeExternalLinkDemo
{
    static void Main()
    {
        try
        {
            // Paths for the main workbook and the external workbook
            string externalPath = "External.xlsx";
            string mainPath = "MainWorkbook.xlsx";

            // -------------------------------------------------
            // 1. Create the external workbook and set a value
            // -------------------------------------------------
            Workbook externalWb = new Workbook();                     // create external workbook
            Worksheet extSheet = externalWb.Worksheets[0];
            extSheet.Name = "Sheet1";
            extSheet.Cells["A1"].PutValue("Hello from external!");   // value to be referenced

            // Save external workbook (overwrite if it already exists)
            externalWb.Save(externalPath, SaveFormat.Xlsx);

            // -------------------------------------------------
            // 2. Create the main workbook
            // -------------------------------------------------
            Workbook mainWb = new Workbook();                         // create main workbook
            Worksheet mainSheet = mainWb.Worksheets[0];
            mainSheet.Name = "MainSheet";

            // -------------------------------------------------
            // 3. Insert a formula that references the external workbook
            // -------------------------------------------------
            // Formula syntax: ='[External.xlsx]Sheet1'!A1
            string externalFormula = $"='[{externalPath}]Sheet1'!A1";
            mainSheet.Cells["B2"].Formula = externalFormula;

            // -------------------------------------------------
            // 4. Add a shape and link it to the cell containing the external reference
            // -------------------------------------------------
            // Add a rectangle shape (row, column, top, left, height, width)
            Shape linkedShape = mainSheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                1,    // upper left row
                0,    // upper left column
                0,    // top offset (pixels)
                0,    // left offset (pixels)
                50,   // height (pixels)
                100); // width (pixels)

            // Set the shape's text (optional)
            linkedShape.Text = "Click to view external value";

            // Configure an internal hyperlink that points to cell B2
            // Use a hash (#) prefix for internal workbook links
            linkedShape.Hyperlink.Address = $"#{mainSheet.Name}!B2";

            // -------------------------------------------------
            // 5. Recalculate formulas and verify the external value is retrieved
            // -------------------------------------------------
            mainWb.CalculateFormula(); // forces calculation, including external reference

            // Retrieve the calculated value from B2
            string retrievedValue = mainSheet.Cells["B2"].StringValue;

            // Simple verification output
            Console.WriteLine("Value in B2 (should match external A1): " + retrievedValue);

            // -------------------------------------------------
            // 6. Save the main workbook
            // -------------------------------------------------
            mainWb.Save(mainPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
