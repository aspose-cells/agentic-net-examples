using System;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerFormulaDemo
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // 2. Define smart markers for data columns (ID, Name, Qty, Price)
            cells["A1"].PutValue("&=$ID");
            cells["B1"].PutValue("&=$Name");
            cells["C1"].PutValue("&=$Qty");
            cells["D1"].PutValue("&=$Price");

            // 3. Place a formula that should be repeated for each inserted row.
            //    The formula references the cells in the same row (C and D).
            cells["E1"].PutValue("=C2*D2"); // This will become =C3*D3, =C4*D4, ...

            // 4. Prepare a data source (DataTable) that matches the smart markers
            DataTable dt = new DataTable("Products");
            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Qty", typeof(int));
            dt.Columns.Add("Price", typeof(double));

            dt.Rows.Add(1, "Apple", 10, 0.5);
            dt.Rows.Add(2, "Banana", 20, 0.3);
            dt.Rows.Add(3, "Cherry", 15, 0.8);

            // 5. Create a WorkbookDesigner, assign the data source, and enable formula repetition
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.RepeatFormulasWithSubtotal = true; // ensures the formula is copied to each generated row
            designer.SetDataSource(dt);

            // 6. Process the smart markers – rows are inserted and the formula adjusts automatically
            designer.Process();

            // 7. Verify the formula in the first data row (optional)
            Console.WriteLine("Formula in E2: " + cells["E2"].Formula); // Expected: =C2*D2
            Console.WriteLine("Formula in E3: " + cells["E3"].Formula); // Expected: =C3*D3
            Console.WriteLine("Formula in E4: " + cells["E4"].Formula); // Expected: =C4*D4

            // 8. Save the resulting workbook
            workbook.Save("SmartMarkerFormulaResult.xlsx");
        }
    }
}