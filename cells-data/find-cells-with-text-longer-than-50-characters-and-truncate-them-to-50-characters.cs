using System;
using Aspose.Cells;

namespace TruncateLongCellText
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file
            string inputPath = "input.xlsx";

            // Load the workbook (default LoadOptions)
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Iterate through all used cells in the worksheet
                foreach (Cell cell in cells)
                {
                    // Process only cells that contain string values
                    if (cell.Type == CellValueType.IsString)
                    {
                        string text = cell.StringValue;
                        if (!string.IsNullOrEmpty(text) && text.Length > 50)
                        {
                            // Truncate to the first 50 characters
                            string truncated = text.Substring(0, 50);
                            cell.PutValue(truncated);
                        }
                    }
                }
            }

            // Save the modified workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
        }
    }
}