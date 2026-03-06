using System;
using Aspose.Cells;

class AddCellWatchesDemo
{
    static void Main()
    {
        // Load an existing XLSX workbook
        LoadOptions loadOptions = new LoadOptions();
        // (optional) loadOptions.ParsingFormulaOnOpen = true;
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add cells to the Formula Watch Window
        int watchIndex1 = sheet.CellWatches.Add("B2"); // watch cell B2
        int watchIndex2 = sheet.CellWatches.Add("C3"); // watch cell C3 (duplicate addition example)

        // Retrieve a watch item (optional demonstration)
        CellWatch watch = sheet.CellWatches[watchIndex1];
        // Example: ensure the watch name matches the cell address
        watch.CellName = "B2";

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}