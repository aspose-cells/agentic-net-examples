using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Sample data: tasks with a "Status" column (0 = pending, 1 = completed)
                cells["A1"].PutValue("Task");
                cells["B1"].PutValue("Status");
                cells["A2"].PutValue("Design UI");
                cells["B2"].PutValue(1); // completed
                cells["A3"].PutValue("Write Code");
                cells["B3"].PutValue(0); // pending
                cells["A4"].PutValue("Test Application");
                cells["B4"].PutValue(1); // completed

                // Create a style that applies a single strikethrough
                Style strikeStyle = workbook.CreateStyle();
                strikeStyle.Font.IsStrikeout = true; // enable strikeout on the font

                // StyleFlag indicating that only the FontStrike property should be applied
                StyleFlag strikeFlag = new StyleFlag();
                strikeFlag.FontStrike = true;

                // Apply the strikethrough style to rows where the status column equals 1 (completed)
                int maxRow = worksheet.Cells.MaxDataRow;
                for (int row = 1; row <= maxRow; row++)
                {
                    if (cells[row, 1].IntValue == 1)
                    {
                        cells.ApplyRowStyle(row, strikeStyle, strikeFlag);
                    }
                }

                // Save the workbook
                string outputPath = "CompletedTasksStrikethrough.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}