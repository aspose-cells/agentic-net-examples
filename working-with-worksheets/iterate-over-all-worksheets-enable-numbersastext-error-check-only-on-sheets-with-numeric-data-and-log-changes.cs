using System;
using Aspose.Cells;

class EnableNumbersAsTextErrorCheck
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            bool hasNumeric = false;

            // Determine the used range of the worksheet
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            // Scan cells for any numeric value
            for (int row = 0; row <= maxRow && !hasNumeric; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    if (cell != null && cell.Type == CellValueType.IsNumeric)
                    {
                        hasNumeric = true;
                        break;
                    }
                }
            }

            // If numeric data exists, enable NumbersStoredAsText error check
            if (hasNumeric)
            {
                // Access the error check option collection for the worksheet
                ErrorCheckOptionCollection options = sheet.ErrorCheckOptions;

                // Add a new error check option
                int optionIndex = options.Add();
                ErrorCheckOption option = options[optionIndex];

                // Enable the "NumberStoredAsText" check
                option.SetErrorCheck(ErrorCheckType.NumberStoredAsText, true);

                // Apply the option to the entire used range of the sheet
                CellArea area = CellArea.CreateCellArea(0, 0, maxRow, maxCol);
                option.AddRange(area);

                // Log the change
                Console.WriteLine($"Enabled NumbersAsText error check on sheet \"{sheet.Name}\".");
            }
            else
            {
                Console.WriteLine($"No numeric data found in sheet \"{sheet.Name}\"; error check not applied.");
            }
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}