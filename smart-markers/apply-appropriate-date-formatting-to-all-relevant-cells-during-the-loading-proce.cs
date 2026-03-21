using System;
using Aspose.Cells;

namespace AsposeCellsDateFormattingDemo
{
    class Program
    {
        static void Main()
        {
            // Input and output file paths
            string inputFile = "input.xlsx";
            string outputFile = "output.xlsx";

            // Create LoadOptions for XLSX format (required by the lifecycle rule)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // Load the workbook using the specified options
            Workbook workbook = new Workbook(inputFile, loadOptions);

            // Access the first worksheet (adjust as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Iterate through all used cells in the worksheet
            for (int row = 0; row <= worksheet.Cells.MaxDataRow; row++)
            {
                for (int col = 0; col <= worksheet.Cells.MaxDataColumn; col++)
                {
                    Cell cell = worksheet.Cells[row, col];

                    // Apply custom date format only to cells that contain DateTime values
                    if (cell.Type == CellValueType.IsDateTime)
                    {
                        // Retrieve the current style, set a custom date format, and apply it back
                        Style style = cell.GetStyle();
                        style.Custom = "yyyy-MM-dd"; // Desired date format
                        cell.SetStyle(style);
                    }
                }
            }

            // Save the modified workbook (required by the lifecycle rule)
            workbook.Save(outputFile);
        }
    }
}