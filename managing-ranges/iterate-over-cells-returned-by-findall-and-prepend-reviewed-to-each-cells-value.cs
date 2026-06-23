using System;
using Aspose.Cells;

namespace AsposeCellsFindAllDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data – you can replace this with loading an existing file
            cells["A1"].PutValue("Item 1");
            cells["B2"].PutValue("Item 2");
            cells["C3"].PutValue(123); // numeric cell – will be left unchanged

            // -----------------------------------------------------------------
            // Iterate over all cells (simulating FindAll) and prepend "Reviewed:"
            // -----------------------------------------------------------------
            foreach (Cell cell in cells)
            {
                // Process only string cells; other types are ignored
                if (cell.Type == CellValueType.IsString)
                {
                    string original = cell.StringValue ?? string.Empty;
                    cell.PutValue("Reviewed:" + original);
                }
            }

            // Save the workbook (replace with your desired path/format)
            workbook.Save("ReviewedOutput.xlsx");
        }
    }
}