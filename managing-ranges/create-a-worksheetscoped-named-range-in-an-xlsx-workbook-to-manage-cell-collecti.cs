using System;
using Aspose.Cells;

namespace AsposeCellsNamedRangeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and give it a friendly name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "DataSheet";

            // Populate sample data
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["B3"].PutValue(15);

            // Add a worksheet‑scoped named range called "MyRange" that refers to A1:B3
            int nameIndex = workbook.Worksheets.Names.Add("MyRange");
            Name myRange = workbook.Worksheets.Names[nameIndex];
            myRange.RefersTo = $"={sheet.Name}!$A$1:$B$3";
            // Set SheetIndex to the current worksheet index for worksheet scope (0‑based)
            myRange.SheetIndex = sheet.Index;

            // Retrieve the range via GetRange and modify the header row (first row) to uppercase
            Aspose.Cells.Range range = myRange.GetRange();
            foreach (Cell cell in range)
            {
                if (cell.Row == range.FirstRow)
                {
                    cell.PutValue(cell.StringValue.ToUpper());
                }
            }

            // Save the workbook
            workbook.Save("WorksheetScopedNamedRange.xlsx");
        }
    }
}