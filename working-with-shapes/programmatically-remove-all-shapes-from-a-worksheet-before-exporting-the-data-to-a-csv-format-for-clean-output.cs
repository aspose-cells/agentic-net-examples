// Title: Remove all shapes from worksheets and export to CSV with Aspose.Cells for .NET
// Description: Loads an Excel workbook, iterates through each worksheet, calls RemoveAllDrawingObjects to delete every shape, chart, picture, and other drawing objects, then saves the cleaned workbook as a CSV file using TxtSaveOptions. Finally disposes the workbook to free resources.
// Keywords: Aspose.Cells RemoveAllDrawingObjects | delete shapes C# | export Excel to CSV Aspose | clean CSV from Excel | remove charts before CSV export | Aspose.Cells drawing objects removal | C# Aspose.Cells CSV conversion
// Common Searches: Aspose.Cells remove all shapes from workbook | How to export Excel to CSV without images using Aspose | C# delete charts before saving as CSV | Remove drawing objects Aspose.Cells .NET | Clean Excel data for CSV conversion
// Developer Intent: Strip every drawing object from each worksheet and then save the workbook as a CSV file.
// Use Cases: Generate pure data CSV reports for systems that cannot handle embedded images or charts. | Batch‑convert multiple Excel files to CSV while guaranteeing only cell values are retained. | Prepare data extracts for migration or analytics pipelines where drawing objects cause parsing errors.
// AI Prompts: Show how to remove only specific shape types (e.g., pictures) before exporting to CSV with Aspose.Cells. | Add robust error handling for missing or corrupt input files when cleaning drawings and saving as CSV. | Create a reusable C# method that takes a workbook path and outputs a CSV file with all drawing objects removed.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeRemoval
{
    // Loads an Excel workbook, iterates through each worksheet, calls RemoveAllDrawingObjects to delete every shape, chart, picture, and other drawing objects, then saves the cleaned workbook as a CSV file using TxtSaveOptions. Finally disposes the workbook to free resources.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your source file)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets and remove every drawing object (shapes, charts, pictures, etc.)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Removes all drawing objects in the current worksheet
                sheet.RemoveAllDrawingObjects();
            }

            // Prepare CSV (TXT) save options – default separator is comma
            TxtSaveOptions csvOptions = new TxtSaveOptions();

            // Export the cleaned workbook to CSV format
            workbook.Save("output.csv", csvOptions);

            // Release resources
            workbook.Dispose();
        }
    }
}
