// Title: Remove a slicer and its pivot connections using Aspose.Cells for .NET (C#)
// Description: Loads an existing workbook, accesses the first worksheet, finds the slicer collection, detaches the target slicer from any linked pivot tables, removes the slicer, and saves the updated file.
// Keywords: Aspose.Cells | C# slicer removal | delete Excel slicer programmatically | pivot table slicer connection | clean up slicer references | remove slicer .NET | Aspose.Cells API slicer
// Common Searches: how to delete a slicer with Aspose.Cells | remove slicer and pivot connections C# | Aspose.Cells delete slicer from worksheet | clean orphaned slicer references Aspose | programmatic slicer removal .NET
// Developer Intent: Programmatically delete an existing slicer and clear all its pivot‑table links to prevent orphaned connections.
// Use Cases: Remove a named slicer (e.g., "RegionSlicer") from a specific sheet while preserving workbook integrity. | Batch‑process a workbook to eliminate every slicer across all worksheets and ensure no residual pivot links remain. | Check a sheet for slicers, safely delete the first one found, and save the modified workbook.
// AI Prompts: Generate C# code with Aspose.Cells that removes a slicer called 'RegionSlicer' and detaches it from all associated pivot tables. | Create a reusable method to delete all slicers in a workbook and automatically clean up their pivot connections. | Write a snippet that iterates through each worksheet, removes any slicers present, and saves the workbook using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;
using Aspose.Cells.Tables;

namespace AsposeCellsSlicerRemoval
{
    // Loads an existing workbook, accesses the first worksheet, finds the slicer collection, detaches the target slicer from any linked pivot tables, removes the slicer, and saves the updated file.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook that contains a table, a pivot table and a slicer
            Workbook workbook = new Workbook("input.xlsx");

            // Assume the slicer is on the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get the slicer collection from the worksheet
            SlicerCollection slicers = sheet.Slicers;

            // Check that there is at least one slicer
            if (slicers.Count == 0)
            {
                Console.WriteLine("No slicers found on the worksheet.");
                return;
            }

            // Retrieve the slicer you want to delete.
            // Here we take the first slicer; you can also use slicers["SlicerName"] or an index you know.
            Slicer slicerToRemove = slicers[0];

            // If the slicer is linked to a PivotTable, remove the connection first.
            // This prevents orphaned references after the slicer is deleted.
            // We iterate through all pivot tables on the worksheet and try to remove the connection.
            foreach (PivotTable pt in sheet.PivotTables)
            {
                try
                {
                    slicerToRemove.RemovePivotConnection(pt);
                }
                catch
                {
                    // If the slicer is not connected to this pivot table, Ignore.
                }
            }

            // Remove the slicer from the collection
            slicers.Remove(slicerToRemove);

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}
