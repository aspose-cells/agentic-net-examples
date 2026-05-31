using System;
using Aspose.Cells;

class TruncateLongText
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Determine the used range limits
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Loop through each cell in the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];

                    // Process only cells that contain string values
                    if (cell.Type == CellValueType.IsString)
                    {
                        string text = cell.StringValue;
                        if (!string.IsNullOrEmpty(text) && text.Length > 50)
                        {
                            // Truncate the text to 50 characters
                            string truncated = text.Substring(0, 50);
                            cell.PutValue(truncated);
                        }
                    }
                }
            }
        }

        // Save the modified workbook
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
    }
}