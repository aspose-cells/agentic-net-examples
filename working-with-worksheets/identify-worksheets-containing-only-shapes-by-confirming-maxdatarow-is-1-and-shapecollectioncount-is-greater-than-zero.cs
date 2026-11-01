// Title: List Excel worksheets that contain only shapes (no cell data) using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that iterates through a workbook and returns the names of worksheets where Cells.MaxDataRow is -1 and the Shapes collection has at least one item. | Extend the sample to export the identified shape‑only worksheet names to a JSON file. | Create a reusable method that accepts a file path and returns a list of worksheet names that contain only drawings, using Aspose.Cells.
// Common Searches: how to find Excel sheets that only have drawings using Aspose.Cells C# | Aspose.Cells detect worksheets with no data but with shapes | C# list worksheets where MaxDataRow equals -1 and Shapes.Count > 0 | filter out shape‑only worksheets from a workbook with Aspose.Cells .NET
// Tags: Aspose.Cells worksheet shape filter | MaxDataRow -1 with Shapes collection detection | C# identify sheets containing only drawings | Excel workbook shape‑only worksheet extraction | filter empty data rows using Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;

// The program loads an Excel workbook, scans each worksheet, and gathers the names of sheets whose Cells.MaxDataRow is -1 (indicating no cell data) while the Shapes collection contains at least one shape, then prints those worksheet names.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // List to store names of worksheets that contain only shapes
        List<string> shapeOnlyWorksheets = new List<string>();

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // MaxDataRow is -1 when there are no data cells
            // ShapeCollection.Count > 0 indicates the presence of shapes
            if (sheet.Cells.MaxDataRow == -1 && sheet.Shapes.Count > 0)
            {
                shapeOnlyWorksheets.Add(sheet.Name);
            }
        }

        // Output the names of worksheets that contain only shapes
        foreach (string name in shapeOnlyWorksheets)
        {
            Console.WriteLine(name);
        }

        // Optional: save the workbook if modifications were made
        // workbook.Save("output.xlsx");
    }
}
