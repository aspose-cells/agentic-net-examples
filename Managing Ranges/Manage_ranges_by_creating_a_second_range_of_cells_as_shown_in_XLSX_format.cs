using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class ManageRangesDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet's cells
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Create the source range (B2:D4) – start at row 1, column 1, 3 rows, 4 columns
        AsposeRange source = cells.CreateRange(1, 1, 3, 4);

        // Fill the source range with sample data
        for (int i = 0; i < source.RowCount; i++)
        {
            for (int j = 0; j < source.ColumnCount; j++)
            {
                source[i, j].PutValue($"R{i}C{j}");
            }
        }

        // Create the destination range (A7:D9) – start at row 6, column 0, same dimensions
        AsposeRange destination = cells.CreateRange(6, 0, 3, 4);

        // Copy only the values from the source range to the destination range
        destination.CopyValue(source);

        // Optionally assign names to the ranges for later reference
        source.Name = "SourceRange";
        destination.Name = "DestinationRange";

        // Save the workbook as an XLSX file
        workbook.Save("ManagedRanges.xlsx");
    }
}