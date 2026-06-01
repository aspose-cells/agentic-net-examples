using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    public class PreserveHeaderFormattingDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate header row (row 0) and some data rows
                sheet.Cells["A1"].PutValue("ID");
                sheet.Cells["B1"].PutValue("Name");
                sheet.Cells["C1"].PutValue("Score");

                for (int i = 2; i <= 5; i++)
                {
                    sheet.Cells[i - 1, 0].PutValue(i - 1);               // ID
                    sheet.Cells[i - 1, 1].PutValue($"Person {i - 1}"); // Name
                    sheet.Cells[i - 1, 2].PutValue((i - 1) * 10);      // Score
                }

                // Create a table that includes the header and data rows
                int tableIndex = sheet.ListObjects.Add(0, 0, 4, 2, true);
                ListObject table = sheet.ListObjects[tableIndex];

                // Apply a distinct style to the header row to demonstrate preservation
                Style headerStyle = workbook.CreateStyle();
                headerStyle.Font.IsBold = true;
                headerStyle.ForegroundColor = Color.LightGray;
                headerStyle.Pattern = BackgroundType.Solid;
                for (int col = 0; col <= 2; col++)
                {
                    Cell headerCell = sheet.Cells[0, col];
                    headerCell.SetStyle(headerStyle);
                }

                // Convert the table to a normal range.
                // TableToRangeOptions does not expose PreserveFormatting in older versions,
                // so we rely on the default behavior which keeps formatting.
                TableToRangeOptions options = new TableToRangeOptions
                {
                    // Limit conversion to the first row (header) only
                    LastRow = 0
                };
                table.ConvertToRange(options);

                // Define output file path
                string outputPath = "PreserveHeaderFormatting.xlsx";

                // Ensure the directory exists (handle possible null directory)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? string.Empty;
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            PreserveHeaderFormattingDemo.Run();
        }
    }
}