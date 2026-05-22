using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    public class TableToRangeRemoveFormatting
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate the worksheet with sample data (5 columns, 10 rows)
                for (int col = 0; col < 5; col++)
                {
                    cells[0, col].PutValue($"Header {col + 1}");
                }
                for (int row = 1; row < 10; row++)
                {
                    for (int col = 0; col < 5; col++)
                    {
                        cells[row, col].PutValue(row * (col + 1));
                    }
                }

                // Add a ListObject (table) covering the populated range
                int tableIndex = worksheet.ListObjects.Add(0, 0, 9, 4, true);
                ListObject table = worksheet.ListObjects[tableIndex];

                // Optionally set a table style (formatting will be removed after conversion)
                table.TableStyleType = TableStyleType.TableStyleMedium2;

                // Create TableToRangeOptions and set LastRow to zero (removes formatting)
                TableToRangeOptions options = new TableToRangeOptions
                {
                    LastRow = 0
                };

                // Convert the table to a normal range using the options
                table.ConvertToRange(options);

                // Save the workbook to verify the result
                string outputPath = "TableToRange_NoFormatting.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            TableToRangeRemoveFormatting.Run();
        }
    }
}