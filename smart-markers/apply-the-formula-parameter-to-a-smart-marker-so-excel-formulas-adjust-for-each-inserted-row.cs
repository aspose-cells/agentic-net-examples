using System;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerFormulaDemo
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // 2. Set up template headers
            cells["A1"].PutValue("ProductID");
            cells["B1"].PutValue("Quantity");
            cells["C1"].PutValue("Total"); // column for formula result

            // 3. Insert a template row that will be expanded by smart markers
            //    Smart markers for data columns
            cells["A2"].PutValue("&=$ProductID");
            cells["B2"].PutValue("&=$Quantity");
            //    Formula that should adjust for each inserted row
            //    Note: use a relative reference (B2) – Aspose.Cells will copy the formula
            //    and automatically adjust the row index for each new row.
            cells["C2"].Formula = "=B2*2";

            // 4. Prepare a data source (DataTable) with sample rows
            DataTable dt = new DataTable("Products");
            dt.Columns.Add("ProductID", typeof(int));
            dt.Columns.Add("Quantity", typeof(int));

            dt.Rows.Add(101, 5);
            dt.Rows.Add(102, 8);
            dt.Rows.Add(103, 12);

            // 5. Create a WorkbookDesigner, assign the data source and process the template
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            // Ensure formulas are repeated for each generated row (default behavior works,
            // but setting this property makes it explicit when subtotal rows are present)
            designer.RepeatFormulasWithSubtotal = true;
            designer.SetDataSource(dt);
            designer.Process(); // lifecycle: process smart markers

            // 6. Verify that formulas have been adjusted and values calculated
            Console.WriteLine("Row\tProductID\tQuantity\tTotal (Formula)");
            for (int row = 1; row <= dt.Rows.Count; row++) // data starts at row index 1 (A2)
            {
                string productId = cells[row, 0].StringValue;
                string quantity = cells[row, 1].StringValue;
                string totalFormula = cells[row, 2].Formula; // should be "=B{row}*2"
                string totalValue = cells[row, 2].StringValue;
                Console.WriteLine($"{row}\t{productId}\t\t{quantity}\t\t{totalFormula} = {totalValue}");
            }

            // 7. Save the workbook (lifecycle: save)
            workbook.Save("SmartMarkerFormulaResult.xlsx");
        }
    }
}