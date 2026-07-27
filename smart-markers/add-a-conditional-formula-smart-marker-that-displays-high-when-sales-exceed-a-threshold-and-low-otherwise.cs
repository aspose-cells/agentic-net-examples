using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerExample
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // 2. Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // 3. Set up headers
            sheet.Cells["A1"].PutValue("Sales");
            sheet.Cells["B1"].PutValue("Status");

            // 4. Insert a smart marker that evaluates the sales value.
            //    The formula displays "High" when Sales > 100, otherwise "Low".
            //    The smart marker syntax uses &= to indicate a formula.
            sheet.Cells["B2"].PutValue("&=IF(Sales>100,\"High\",\"Low\")");

            // 5. Prepare a data source (DataTable) with a Sales column.
            DataTable salesTable = new DataTable("Data");
            salesTable.Columns.Add("Sales", typeof(double));
            // Example rows – in a real scenario these would come from your data source.
            salesTable.Rows.Add(120); // Should display "High"
            salesTable.Rows.Add(80);  // Should display "Low"

            // 6. Create a WorkbookDesigner, assign the workbook and the data source.
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(salesTable);

            // 7. Process the smart markers (rule: Process()).
            designer.Process();

            // 8. Save the result (lifecycle rule)
            workbook.Save("SmartMarkerConditionalOutput.xlsx");
        }
    }
}