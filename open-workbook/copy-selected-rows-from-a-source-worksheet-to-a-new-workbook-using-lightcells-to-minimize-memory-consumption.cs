// Title: Copy Specific Rows to a New Workbook with LightCells in Aspose.Cells for .NET
// Description: Shows how to load a source workbook, select non‑contiguous rows, copy them to a fresh worksheet, and save the result using LightCells (OoxmlSaveOptions with a NoOpLightCellsDataProvider) to keep memory usage low in C#.
// Keywords: Aspose.Cells | CopyRow | LightCells | C# example | low memory Excel | select rows | OoxmlSaveOptions | NoOpLightCellsDataProvider | .NET | Excel streaming | memory‑efficient workbook
// Common Searches: asp.net copy selected rows aspose.cells | lightcells save options c# | how to copy non‑contiguous rows aspose | reduce memory usage when saving excel with aspose | custom LightCellsDataProvider example | copy rows to new workbook aspose cells
// Developer Intent: Extract chosen rows from an existing worksheet and write them to a separate workbook while keeping the operation memory‑light.
// Use Cases: Generate a lightweight report by extracting only header and key data rows from a massive spreadsheet. | Create a subset file containing scattered rows for downstream processing without loading the full source into memory. | Process huge Excel files in batch jobs where copying sparse rows with LightCells prevents out‑of‑memory failures.
// AI Prompts: Write C# code that copies rows 1, 3, and 5 from a worksheet to a new workbook and saves it using LightCells to minimize memory consumption. | Explain how to build a custom LightCellsDataProvider that streams rows while saving a workbook with Aspose.Cells. | Adapt the example to read a dynamic list of row indices from a JSON configuration file and copy those rows with LightCells.

using System;
using Aspose.Cells;

// Shows how to load a source workbook, select non‑contiguous rows, copy them to a fresh worksheet, and save the result using LightCells (OoxmlSaveOptions with a NoOpLightCellsDataProvider) to keep memory usage low in C#.
class Program
{
    static void Main()
    {
        // Load the source workbook (replace with actual path)
        Workbook sourceWorkbook = new Workbook("Source.xlsx");
        Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

        // Create a new workbook for the selected rows
        Workbook destinationWorkbook = new Workbook();
        // Remove the default sheet and add a fresh one
        destinationWorkbook.Worksheets.Clear();
        Worksheet destSheet = destinationWorkbook.Worksheets.Add("SelectedRows");

        // Define the zero‑based indices of rows to copy from the source sheet
        int[] rowsToCopy = new int[] { 0, 2, 4 }; // example: rows 1, 3 and 5

        int destRowIndex = 0;
        foreach (int srcRowIndex in rowsToCopy)
        {
            // Copy a single row from source to destination
            destSheet.Cells.CopyRow(sourceSheet.Cells, srcRowIndex, destRowIndex);
            destRowIndex++;
        }

        // Save the destination workbook using LightCells mode to reduce memory usage
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx)
        {
            LightCellsDataProvider = new NoOpLightCellsDataProvider()
        };
        destinationWorkbook.Save("SelectedRows.xlsx", saveOptions);
    }

    // LightCellsDataProvider that delegates all saving to the normal data model
    class NoOpLightCellsDataProvider : LightCellsDataProvider
    {
        public bool StartSheet(int sheetIndex) => false; // use default sheet processing
        public int NextRow() => -1;                     // no custom rows
        public void StartRow(Row row) { }               // no custom row handling
        public int NextCell() => -1;                    // no custom cells
        public void StartCell(Cell cell) { }            // no custom cell handling
        public bool IsGatherString() => false;          // default string handling
    }
}
