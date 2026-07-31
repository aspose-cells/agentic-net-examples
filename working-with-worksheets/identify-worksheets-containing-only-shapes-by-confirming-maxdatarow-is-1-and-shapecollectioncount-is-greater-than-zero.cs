// Title: Detect Excel worksheets that contain only shapes with Aspose.Cells for .NET
// Description: This C# example loads an Excel file, scans each worksheet, and uses `Cells.MaxDataRow == -1` together with `Shapes.Count > 0` to identify sheets that have no cell data but include one or more drawing objects. The names of such worksheets are written to the console, and the workbook is saved unchanged.
// Keywords: Aspose.Cells C# shape detection | MaxDataRow -1 worksheet | Shapes.Count Excel | identify drawing‑only sheets | empty data worksheet Aspose | Excel worksheet without cells but with shapes | C# Aspose.Cells example
// Common Searches: how to find worksheets with only drawings using Aspose.Cells | C# check if Excel sheet has no data rows but contains shapes | Aspose.Cells detect shape‑only worksheets | list Excel sheets that are empty of cells but have graphics
// Developer Intent: Locate worksheets that lack any cell content yet contain at least one shape object.
// Use Cases: Create an audit report of drawing‑only sheets before publishing a workbook. | Skip non‑data worksheets during bulk export or conversion processes. | Flag or remove shape‑only tabs to reduce file size and improve performance.
// AI Prompts: Generate C# code with Aspose.Cells to delete all worksheets that contain only shapes. | Show how to extract every shape from data‑less worksheets and save them as separate image files. | Explain how to extend the sample so that charts are also treated as shapes when detecting shape‑only sheets.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This C# example loads an Excel file, scans each worksheet, and uses `Cells.MaxDataRow == -1` together with `Shapes.Count > 0` to identify sheets that have no cell data but include one or more drawing objects. The names of such worksheets are written to the console, and the workbook is saved unchanged.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace the path with your file)
        var loadOptions = new LoadOptions();
        // Optional: ignore overlapping useless shapes during load
        loadOptions.IgnoreUselessShapes = true;
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // MaxDataRow == -1 means the worksheet has no cell data
            bool hasNoData = sheet.Cells.MaxDataRow == -1;

            // ShapeCollection.Count > 0 means the worksheet contains at least one shape
            bool hasShapes = sheet.Shapes.Count > 0;

            // Identify worksheets that contain only shapes (no data rows)
            if (hasNoData && hasShapes)
            {
                Console.WriteLine($"Worksheet \"{sheet.Name}\" contains only shapes.");
            }
        }

        // Save the workbook (no modifications made, just demonstrating save lifecycle)
        workbook.Save("output.xlsx");
    }
}
