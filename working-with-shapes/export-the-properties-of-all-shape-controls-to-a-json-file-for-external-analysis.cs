// Title: Export all shape control properties to JSON using Aspose.Cells for .NET
// Description: Creates a workbook, adds sample shapes, records each shape's worksheet name, shape name, type, hidden flag, control‑data length and Base64‑encoded control data on a "ShapeInfo" sheet, then saves that sheet as a formatted JSON file with JsonSaveOptions.
// Keywords: Aspose.Cells | C# | export shape properties | JSON output | ControlData Base64 | JsonSaveOptions | shape metadata extraction | Excel shape export
// Common Searches: Aspose.Cells export shape properties to JSON | Get ControlData of shapes Aspose.Cells .NET | Save specific worksheet as JSON Aspose.Cells | Iterate workbook shapes and export attributes | Base64 encode shape control data Aspose
// Developer Intent: Extract every shape's attributes and control data from a workbook and write them to a JSON file.
// Use Cases: Audit all form controls in an Excel file for compliance reporting. | Send shape metadata to a web API by converting ControlData to Base64 strings. | Create a lightweight JSON snapshot of shape layout for version control or documentation.
// AI Prompts: Generate C# code with Aspose.Cells that reads an existing workbook, extracts each shape's Name, Type, IsHidden, ControlData length and Base64 value, and writes the data to a JSON file. | Adapt the example to export only rectangle shapes and include their top‑left coordinates in the JSON output. | Explain how to decode the Base64 ControlData from the produced JSON back to its original binary form in C#.

using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeControlExport
{
    // Creates a workbook, adds sample shapes, records each shape's worksheet name, shape name, type, hidden flag, control‑data length and Base64‑encoded control data on a "ShapeInfo" sheet, then saves that sheet as a formatted JSON file with JsonSaveOptions.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "DemoSheet";

            // Add sample shapes (for demonstration purposes)
            Shape rect = sheet.Shapes.AddRectangle(1, 1, 100, 100, 0, 0);
            rect.Name = "Rectangle1";
            rect.IsHidden = false;

            Shape oval = sheet.Shapes.AddOval(2, 2, 80, 80, 0, 0);
            oval.Name = "Oval1";
            oval.IsHidden = true;

            // Create a worksheet to hold shape property data
            int infoSheetIdx = workbook.Worksheets.Add();
            Worksheet infoSheet = workbook.Worksheets[infoSheetIdx];
            infoSheet.Name = "ShapeInfo";

            // Write header row
            infoSheet.Cells["A1"].PutValue("Worksheet");
            infoSheet.Cells["B1"].PutValue("ShapeName");
            infoSheet.Cells["C1"].PutValue("ShapeType");
            infoSheet.Cells["D1"].PutValue("IsHidden");
            infoSheet.Cells["E1"].PutValue("ControlDataLength");
            infoSheet.Cells["F1"].PutValue("ControlDataBase64");

            int row = 1; // zero‑based index; row 1 is the second row (after header)

            // Iterate through all worksheets and their shapes
            foreach (Worksheet ws in workbook.Worksheets)
            {
                foreach (Shape shp in ws.Shapes)
                {
                    // Basic properties
                    infoSheet.Cells[row, 0].PutValue(ws.Name);
                    infoSheet.Cells[row, 1].PutValue(shp.Name);
                    infoSheet.Cells[row, 2].PutValue(shp.Type.ToString());
                    infoSheet.Cells[row, 3].PutValue(shp.IsHidden);

                    // ControlData handling
                    byte[] ctrlData = shp.ControlData;
                    if (ctrlData != null && ctrlData.Length > 0)
                    {
                        infoSheet.Cells[row, 4].PutValue(ctrlData.Length);
                        infoSheet.Cells[row, 5].PutValue(Convert.ToBase64String(ctrlData));
                    }
                    else
                    {
                        infoSheet.Cells[row, 4].PutValue(0);
                        infoSheet.Cells[row, 5].PutValue(string.Empty);
                    }

                    row++;
                }
            }

            // Configure JSON save options to export only the ShapeInfo sheet
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                SheetIndexes = new int[] { infoSheetIdx },
                ExportEmptyCells = true,
                HasHeaderRow = true,
                ExportAsString = true,
                Indent = "  "
            };

            // Save the workbook as a JSON file containing shape properties
            workbook.Save("ShapeProperties.json", jsonOptions);
        }
    }
}
