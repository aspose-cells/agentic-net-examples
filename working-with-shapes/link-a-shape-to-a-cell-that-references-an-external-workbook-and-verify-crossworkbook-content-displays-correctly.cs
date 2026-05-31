using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExternalShapeLinkDemo
{
    class Program
    {
        static void Main()
        {
            // ---------- Step 1: Create external workbook ----------
            Workbook externalWb = new Workbook();
            Worksheet externalWs = externalWb.Worksheets[0];
            externalWs.Name = "Sheet1";
            externalWs.Cells["A1"].PutValue("External Data");
            string externalFileName = "ExternalWorkbook.xlsx";
            externalWb.Save(externalFileName);
            
            // ---------- Step 2: Create main workbook ----------
            Workbook mainWb = new Workbook();
            Worksheet mainWs = mainWb.Worksheets[0];
            mainWs.Name = "MainSheet";

            // ---------- Step 3: Add external link ----------
            // Add the external workbook to the external links collection
            int linkIndex = mainWb.Worksheets.ExternalLinks.Add(externalFileName, new string[] { "Sheet1" });

            // ---------- Step 4: Set formula that references the external workbook ----------
            // Cell B1 will pull the value from external workbook's A1
            mainWs.Cells["B1"].Formula = $"='[{externalFileName}]Sheet1'!A1";

            // ---------- Step 5: Add a shape and link it to the cell ----------
            // Add a rectangle shape
            Shape shape = mainWs.Shapes.AddRectangle(2, 2, 100, 50, 0, 0);
            // Link the shape to cell B1
            shape.LinkedCell = "$B$1";

            // ---------- Step 6: Update linked data source ----------
            // Load the external workbook again (could reuse externalWb, but demonstrating the API)
            Workbook externalWbForUpdate = new Workbook(externalFileName);
            mainWb.UpdateLinkedDataSource(new Workbook[] { externalWbForUpdate });

            // ---------- Step 7: Calculate formulas ----------
            mainWb.CalculateFormula();

            // ---------- Step 8: Verify that the shape reflects the external data ----------
            // The linked cell B1 should now contain "External Data"
            string linkedCellValue = mainWs.Cells["B1"].StringValue;
            Console.WriteLine($"Value in linked cell B1: {linkedCellValue}");
            Console.WriteLine($"Shape's LinkedCell property: {shape.LinkedCell}");

            // ---------- Step 9: Save the main workbook ----------
            string outputFile = "MainWorkbook_WithLinkedShape.xlsx";
            mainWb.Save(outputFile);
            Console.WriteLine($"Main workbook saved as {outputFile}");

            // Clean up
            externalWb.Dispose();
            externalWbForUpdate.Dispose();
            mainWb.Dispose();
        }
    }
}