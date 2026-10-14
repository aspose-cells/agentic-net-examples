// Title: Clone an Excel workbook, assign new GUID‑based TabIds, remove empty worksheets, and save the optimized copy with Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells to duplicate a workbook via a MemoryStream, generate a unique GUID‑derived TabId for each worksheet, delete any sheet that has no data cells, and write the cleaned workbook to a new XLSX file. | Create a copy of an existing Excel file, reassign fresh TabId values to all worksheets, strip out blank sheets, and save the resulting workbook using the Aspose.Cells .NET API.
// Common Searches: aspnet clone workbook and reset worksheet TabId using Aspose.Cells | remove blank worksheets from an Excel file with Aspose.Cells C# | how to assign new TabId to each sheet after copying workbook in Aspose.Cells | save duplicated workbook to memory stream then to file Aspose.Cells .NET | optimize Excel workbook by deleting empty sheets with Aspose.Cells
// Tags: workbook cloning with in‑memory stream Aspose.Cells | worksheet TabId reset C# | empty worksheet removal Aspose.Cells | optimized Excel save .NET | Aspose.Cells workbook cleanup workflow

using System;
using System.IO;
using Aspose.Cells;

// Loads an existing XLSX file, clones it via a MemoryStream, assigns a new GUID‑based TabId to each worksheet, removes worksheets that contain no data, and saves the cleaned workbook as a new XLSX file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the original workbook (create/load rule)
        string originalPath = "original.xlsx";
        Workbook originalWorkbook = new Workbook(originalPath);

        // Clone the workbook by saving to a memory stream and loading it back
        Workbook clonedWorkbook;
        using (MemoryStream ms = new MemoryStream())
        {
            originalWorkbook.Save(ms, SaveFormat.Xlsx);
            ms.Position = 0;
            clonedWorkbook = new Workbook(ms);
        }

        // Assign new TabIds to each worksheet
        for (int i = 0; i < clonedWorkbook.Worksheets.Count; i++)
        {
            Worksheet ws = clonedWorkbook.Worksheets[i];
            // Generate a new unique identifier for the tab
            ws.TabId = Guid.NewGuid().GetHashCode();
        }

        // Remove empty worksheets (no data cells)
        for (int i = clonedWorkbook.Worksheets.Count - 1; i >= 0; i--)
        {
            Worksheet ws = clonedWorkbook.Worksheets[i];
            // MaxDataRow/MaxDataColumn are -1 when the sheet is empty
            if (ws.Cells.MaxDataRow < 0 || ws.Cells.MaxDataColumn < 0)
            {
                clonedWorkbook.Worksheets.RemoveAt(i);
            }
        }

        // Save the optimized copy (save rule)
        string outputPath = "optimized_copy.xlsx";
        clonedWorkbook.Save(outputPath, SaveFormat.Xlsx);
    }
}
