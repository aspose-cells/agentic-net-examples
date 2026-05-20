using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsThemeDemo
{
    public class SetBuiltinThemeAndApplyToAllCells
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add sample data to demonstrate the effect
                Worksheet ws = workbook.Worksheets[0];
                ws.Cells["A1"].PutValue("Header");
                ws.Cells["A2"].PutValue(123);
                ws.Cells["B1"].PutValue(DateTime.Now);
                ws.Cells["B2"].PutValue("Sample Text");

                // Create a built‑in style (e.g., Good) and set it as the workbook's default style
                Style builtinStyle = workbook.CreateBuiltinStyle(BuiltinStyleType.Good);
                workbook.DefaultStyle = builtinStyle;

                // Apply the default style to every used cell in all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;
                    int maxRow = cells.MaxDataRow;
                    int maxCol = cells.MaxDataColumn;

                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            // Apply only to cells that contain data or have a style
                            if (cells[row, col].GetStyle() != null)
                            {
                                cells[row, col].SetStyle(workbook.DefaultStyle);
                            }
                        }
                    }
                }

                // Define output file path
                string outputPath = "WorkbookWithBuiltinTheme.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
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
            SetBuiltinThemeAndApplyToAllCells.Run();
        }
    }
}