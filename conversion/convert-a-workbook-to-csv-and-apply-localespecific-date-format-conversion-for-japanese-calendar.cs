using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class WorkbookToCsvJapaneseCalendar
    {
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                // Path to the source Excel workbook
                string sourcePath = "input.xlsx";

                // Verify the source file exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the workbook (using default LoadOptions)
                Workbook workbook = new Workbook(sourcePath);

                // Set the workbook region to Japan to enable Japanese calendar formatting
                workbook.Settings.Region = CountryCode.Japan;

                // Apply a Japanese date format to all cells that contain DateTime values
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;
                    // Iterate through the used range of the worksheet
                    for (int row = 0; row <= cells.MaxDataRow; row++)
                    {
                        for (int col = 0; col <= cells.MaxDataColumn; col++)
                        {
                            Cell cell = cells[row, col];
                            if (cell.Type == CellValueType.IsDateTime)
                            {
                                // Retrieve the cell style
                                Style style = cell.GetStyle();
                                // Japanese calendar custom format (example: 2020年9月15日)
                                style.Custom = "[$-F800]yyyy年m月d日";
                                cell.SetStyle(style);
                            }
                        }
                    }
                }

                // Save the workbook as CSV
                string destinationPath = "output.csv";
                workbook.Save(destinationPath, SaveFormat.Csv);

                Console.WriteLine($"Workbook converted to CSV with Japanese date formatting: {destinationPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}