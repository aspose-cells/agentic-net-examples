using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

public class Program
{
    public static void Main()
    {
        Run();
    }

    public static void Run()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("source.xlsx");

        // Access the first worksheet and its cells collection
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Define the source range (e.g., A1:C3)
        AsposeRange sourceRange = cells.CreateRange("A1", "C3");

        // Determine the size of the source range
        int rowCount = sourceRange.RowCount;
        int columnCount = sourceRange.ColumnCount;

        // Define the destination range starting at column E (index 4) with the same dimensions
        AsposeRange destinationRange = cells.CreateRange(0, 4, rowCount, columnCount);

        // Configure paste options to copy everything (values, formulas, formats, etc.)
        PasteOptions pasteOptions = new PasteOptions
        {
            PasteType = PasteType.All
        };

        // Copy the source range into the destination range
        destinationRange.Copy(sourceRange, pasteOptions);

        // Save the modified workbook to a new file
        workbook.Save("output.xlsx");
    }
}