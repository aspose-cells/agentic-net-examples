// Title: Export Excel Shape Control Data to JSON with Aspose.Cells for .NET
// Description: Loads an Excel workbook, walks through every worksheet and each shape, captures the shape's name, type, and any ControlData (encoded as Base64), writes these records to a temporary sheet with headers, and saves the sheet as a JSON file using Aspose.Cells JsonSaveOptions (HasHeaderRow, ExportEmptyCells, ExportStylePool).
// Keywords: Aspose.Cells | C# export shapes to JSON | Excel shape control data | Base64 shape data | JsonSaveOptions example | .NET workbook shape extraction | shape metadata JSON
// Common Searches: export shape properties to json aspose.cells | retrieve controldata from excel shapes c# | convert shape data to base64 aspose | jsonsaveoptions hasheaderrow example | list all shapes in workbook aspose.cells
// Developer Intent: Extract the name, type, and ControlData of every shape in an Excel file and serialize the collection to a JSON document.
// Use Cases: Create an inventory of form controls for downstream analytics. | Generate a JSON manifest for automated UI testing or reporting. | Archive shape metadata for version control, auditing, or migration.
// AI Prompts: Write C# code that opens an Excel file with Aspose.Cells, iterates all shapes, captures Name, Type, and ControlData as a Base64 string, and saves the result to a JSON file using JsonSaveOptions. | Explain how HasHeaderRow, ExportEmptyCells, and ExportStylePool affect the JSON output when exporting shape information with Aspose.Cells.

using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Json;

namespace AsposeCellsShapeExport
{
    // Loads an Excel workbook, walks through every worksheet and each shape, captures the shape's name, type, and any ControlData (encoded as Base64), writes these records to a temporary sheet with headers, and saves the sheet as a JSON file using Aspose.Cells JsonSaveOptions (HasHeaderRow, ExportEmptyCells, ExportStylePool).
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file that contains shapes
            string sourcePath = "input.xlsx";

            // Load the workbook that contains the shapes
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Create a new workbook that will hold the shape information
            Workbook exportWorkbook = new Workbook();
            Worksheet exportSheet = exportWorkbook.Worksheets[0];

            // Write header row
            exportSheet.Cells["A1"].PutValue("Worksheet");
            exportSheet.Cells["B1"].PutValue("ShapeName");
            exportSheet.Cells["C1"].PutValue("ShapeType");
            exportSheet.Cells["D1"].PutValue("ControlDataBase64");

            int currentRow = 1; // zero‑based index; row 1 is the second row (after header)

            // Iterate through all worksheets in the source workbook
            foreach (Worksheet ws in sourceWorkbook.Worksheets)
            {
                // Iterate through all shapes in the current worksheet
                foreach (Shape shape in ws.Shapes)
                {
                    // Retrieve shape name (if not set, use empty string)
                    string shapeName = shape.Name ?? string.Empty;

                    // Retrieve shape type as string
                    string shapeType = shape.Type.ToString();

                    // Retrieve control data and convert to Base64 string (null if no data)
                    string controlDataBase64 = null;
                    byte[] controlData = shape.ControlData;
                    if (controlData != null && controlData.Length > 0)
                    {
                        controlDataBase64 = Convert.ToBase64String(controlData);
                    }

                    // Populate the export sheet
                    exportSheet.Cells[currentRow, 0].PutValue(ws.Name);          // Worksheet name
                    exportSheet.Cells[currentRow, 1].PutValue(shapeName);       // Shape name
                    exportSheet.Cells[currentRow, 2].PutValue(shapeType);       // Shape type
                    exportSheet.Cells[currentRow, 3].PutValue(controlDataBase64); // Control data

                    currentRow++;
                }
            }

            // Configure JSON save options
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                // Export the header row as column names
                HasHeaderRow = true,
                // Export empty cells as null to keep structure consistent
                ExportEmptyCells = true,
                // Do not export styles (not needed for analysis)
                ExportStylePool = false
            };

            // Save the workbook as a JSON file using Aspose.Cells saving mechanism
            string outputPath = "shapes_export.json";
            exportWorkbook.Save(outputPath, jsonOptions);

            Console.WriteLine($"Shape information exported to JSON file: {outputPath}");
        }
    }
}
