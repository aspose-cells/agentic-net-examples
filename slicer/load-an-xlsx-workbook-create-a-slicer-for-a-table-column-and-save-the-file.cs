// Title: Add a slicer to the first column of a table in an existing XLSX workbook and save with Aspose.Cells for .NET (C#)
// AI Prompts: Load Input.xlsx, locate the first ListObject, add a slicer for its first column at cell E1, assign the caption "My Slicer", and save the workbook as Output.xlsx using Aspose.Cells in C#. | Using Aspose.Cells for .NET, programmatically create a slicer for a table column, position it at a target cell, customize its caption, and export the updated XLSX file.
// Common Searches: Aspose.Cells .NET place slicer on Excel table column | C# code to create slicer for first column of first table in XLSX | how to set slicer caption with Aspose.Cells in C# | save workbook after adding slicer using Aspose.Cells | Aspose.Cells example for slicer positioning at specific cell
// Tags: Aspose.Cells ListObject slicer creation | Aspose.Cells slicer placement at E1 | Excel table slicer caption customization | save modified XLSX with Aspose.Cells | programmatic slicer creation .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Slicers;

// The example loads Input.xlsx, verifies a table exists, adds a slicer for the first column of the first table at cell E1, sets the slicer's caption to "My Slicer", and saves the updated workbook as Output.xlsx.
class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("Input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure the worksheet contains at least one table (ListObject)
        if (worksheet.ListObjects.Count == 0)
        {
            Console.WriteLine("No table found in the worksheet.");
            return;
        }

        // Get the first table in the worksheet
        ListObject table = worksheet.ListObjects[0];

        // Index of the column in the table for which the slicer will be created (zero‑based)
        int columnIndex = 0; // e.g., first column

        // Destination cell name for the top‑left corner of the slicer
        string destCellName = "E1";

        // Add a slicer using the matching rule: Add(ListObject, int, string)
        int slicerIndex = worksheet.Slicers.Add(table, columnIndex, destCellName);

        // Retrieve the slicer object to customize it (optional)
        Slicer slicer = worksheet.Slicers[slicerIndex];
        slicer.Caption = "My Slicer";

        // Save the modified workbook
        workbook.Save("Output.xlsx");
    }
}
