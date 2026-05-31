using System;
using System.Text;
using Aspose.Cells;

namespace MergedCellsSummary
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // Sample data: create a couple of merged ranges with string values
            // ------------------------------------------------------------
            // Merge A1:B2 and put a string in the merged cell
            sheet.Cells.Merge(0, 0, 2, 2);               // A1:B2
            sheet.Cells["A1"].PutValue("FirstPart");

            // Merge A4:B5 and put another string
            sheet.Cells.Merge(3, 0, 2, 2);               // A4:B5
            sheet.Cells["A4"].PutValue("SecondPart");

            // ------------------------------------------------------------
            // Retrieve raw string values from all merged cells,
            // concatenate them, and write the result to a summary cell (C1)
            // ------------------------------------------------------------
            StringBuilder concatenated = new StringBuilder();

            // Get all merged areas in the worksheet
            CellArea[] mergedAreas = sheet.Cells.GetMergedAreas();

            foreach (CellArea area in mergedAreas)
            {
                // Iterate through each cell inside the merged area
                for (int row = area.StartRow; row <= area.EndRow; row++)
                {
                    for (int col = area.StartColumn; col <= area.EndColumn; col++)
                    {
                        Cell cell = sheet.Cells[row, col];

                        // Check if the cell contains a string value
                        if (cell.Type == CellValueType.IsString)
                        {
                            string value = cell.StringValue;
                            if (!string.IsNullOrEmpty(value))
                            {
                                concatenated.Append(value);
                                concatenated.Append(" "); // separator
                            }
                        }
                    }
                }
            }

            // Trim the trailing space and store the result in cell C1 (row 0, column 2)
            sheet.Cells[0, 2].PutValue(concatenated.ToString().Trim());

            // Save the workbook to a file
            workbook.Save("MergedCellsSummary.xlsx");
        }
    }
}