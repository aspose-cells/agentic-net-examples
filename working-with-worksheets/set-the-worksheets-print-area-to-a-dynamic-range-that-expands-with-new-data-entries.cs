using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace DynamicPrintAreaDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate initial data (header + 5 rows)
                sheet.Cells["A1"].PutValue("ID");
                sheet.Cells["B1"].PutValue("Value");
                for (int i = 2; i <= 6; i++)
                {
                    sheet.Cells[$"A{i}"].PutValue(i - 1);
                    sheet.Cells[$"B{i}"].PutValue((i - 1) * 10);
                }

                // Set the print area to the current used range
                SetPrintAreaToUsedRange(sheet);

                // Simulate adding more data later
                for (int i = 7; i <= 12; i++)
                {
                    sheet.Cells[$"A{i}"].PutValue(i - 1);
                    sheet.Cells[$"B{i}"].PutValue((i - 1) * 10);
                }

                // Update the print area so it expands to include the new rows
                SetPrintAreaToUsedRange(sheet);

                // Save the workbook
                string outputPath = "DynamicPrintArea.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        // Helper method that sets the worksheet's print area to the maximal display range
        static void SetPrintAreaToUsedRange(Worksheet sheet)
        {
            // MaxDisplayRange returns a Range that covers all data, merged cells and shapes
            AsposeRange maxRange = sheet.Cells.MaxDisplayRange;
            if (maxRange == null) return; // empty sheet

            int startRow = maxRange.FirstRow;
            int startCol = maxRange.FirstColumn;
            int endRow = startRow + maxRange.RowCount - 1;
            int endCol = startCol + maxRange.ColumnCount - 1;

            // Convert cell indices to Excel addresses (e.g., A1, B5)
            string startAddr = CellsHelper.CellIndexToName(startRow, startCol);
            string endAddr = CellsHelper.CellIndexToName(endRow, endCol);

            // Assign the address range to the PrintArea property
            sheet.PageSetup.PrintArea = $"{startAddr}:{endAddr}";
        }
    }
}