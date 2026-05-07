using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        // Load the demonstration workbook (replace with actual file path)
        string inputFile = "DemoFile.xlsx";
        Workbook workbook;

        if (File.Exists(inputFile))
        {
            workbook = new Workbook(inputFile);
        }
        else
        {
            // Create a new workbook if the file does not exist
            workbook = new Workbook();
        }

        // Access the first worksheet and its cells collection
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // ------------------------------------------------------------
        // 1. Create a range covering cells A2:C5 using the string overload
        // ------------------------------------------------------------
        AsposeRange dataRange = cells.CreateRange("A2", "C5");
        dataRange.Name = "DataRange"; // optional naming for later reference

        // ------------------------------------------------------------
        // 2. Clear both contents and formatting of the same area
        //    using ClearRange with a CellArea object
        // ------------------------------------------------------------
        CellArea clearArea = CellArea.CreateCellArea("A2", "C5");
        cells.ClearRange(clearArea);

        // ------------------------------------------------------------
        // 3. Populate the cleared range with new sample values
        // ------------------------------------------------------------
        for (int r = 0; r < dataRange.RowCount; r++)
        {
            for (int c = 0; c < dataRange.ColumnCount; c++)
            {
                // Example value: "R2C1", "R2C2", ...
                dataRange[r, c].PutValue($"R{r + 2}C{c + 1}");
            }
        }

        // ------------------------------------------------------------
        // 4. Merge the first row of the range (A2:C2) into a single cell
        //    Merge(startRow, startColumn, totalRows, totalColumns)
        //    Note: zero‑based indices, so A2 is row 1, column 0
        // ------------------------------------------------------------
        cells.Merge(1, 0, 1, 3); // merges A2:C2

        // ------------------------------------------------------------
        // 5. Save the modified workbook to a new file
        // ------------------------------------------------------------
        string outputFile = "DemoFile_Modified.xlsx";
        workbook.Save(outputFile);
    }
}