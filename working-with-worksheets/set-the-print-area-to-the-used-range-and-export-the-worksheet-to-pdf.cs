using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data
            worksheet.Cells["A1"].PutValue("Name");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["A2"].PutValue("Item1");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["A3"].PutValue("Item2");
            worksheet.Cells["B3"].PutValue(20);

            // Determine the used range of the worksheet
            AsposeRange usedRange = worksheet.Cells.MaxDisplayRange;

            // Calculate start and end cell indices
            int startRow = usedRange.FirstRow;
            int startColumn = usedRange.FirstColumn;
            int endRow = startRow + usedRange.RowCount - 1;
            int endColumn = startColumn + usedRange.ColumnCount - 1;

            // Convert indices to A1 style addresses
            string startCell = CellsHelper.CellIndexToName(startRow, startColumn);
            string endCell = CellsHelper.CellIndexToName(endRow, endColumn);

            // Set the print area to the used range
            worksheet.PageSetup.PrintArea = $"{startCell}:{endCell}";

            // Export the worksheet to PDF
            string outputPath = "UsedRangePrintArea.pdf";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}