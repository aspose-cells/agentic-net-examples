using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class SetPrintAreaFromMaxDisplayRange
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Sample data to create a display range
            worksheet.Cells["A1"].PutValue("Header1");
            worksheet.Cells["B1"].PutValue("Header2");
            worksheet.Cells["A2"].PutValue(100);
            worksheet.Cells["B2"].PutValue(200);
            worksheet.Cells["A3"].PutValue(300);
            worksheet.Cells["B3"].PutValue(400);

            // Retrieve the maximum display range (includes data, merged cells, shapes)
            AsposeRange maxRange = worksheet.Cells.MaxDisplayRange;

            if (maxRange != null)
            {
                // Determine the start cell name
                string startCell = worksheet.Cells[maxRange.FirstRow, maxRange.FirstColumn].Name;

                // Determine the end cell name
                int endRow = maxRange.FirstRow + maxRange.RowCount - 1;
                int endCol = maxRange.FirstColumn + maxRange.ColumnCount - 1;
                string endCell = worksheet.Cells[endRow, endCol].Name;

                // Set the print area to the max display range
                worksheet.PageSetup.PrintArea = $"{startCell}:{endCell}";
            }

            // Save the workbook
            string outputPath = "PrintAreaFromMaxDisplayRange.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}