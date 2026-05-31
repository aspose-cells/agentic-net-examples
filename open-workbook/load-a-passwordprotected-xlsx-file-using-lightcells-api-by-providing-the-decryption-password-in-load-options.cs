using System;
using System.IO;
using Aspose.Cells;

namespace LoadPasswordProtectedFileApp
{
    class LoadPasswordProtectedFile
    {
        static void Main()
        {
            string filePath = "protected.xlsx";
            string password = "test";

            try
            {
                // Verify that the file exists before attempting to load it
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    return;
                }

                // Configure load options with the workbook password
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    Password = password
                };

                // Load the password‑protected workbook
                Workbook workbook = new Workbook(filePath, loadOptions);

                // Iterate through each worksheet and its used cells
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;
                    int maxRow = cells.MaxDataRow;
                    int maxCol = cells.MaxDataColumn;

                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = cells[row, col];
                            if (cell.Value != null)
                            {
                                Console.WriteLine($"Sheet {sheet.Index}, Cell [{row}, {col}] = {cell.Value}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}