using System;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerCallbackDemo
{
    // Custom callback that modifies each cell after a smart marker is processed
    public class MySmartMarkerCallback : ISmartMarkerCallBack
    {
        private readonly Workbook _workbook;

        public MySmartMarkerCallback(Workbook workbook)
        {
            _workbook = workbook;
        }

        // This method is called for every smart marker cell during processing
        public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
        {
            // Get the cell that has just been populated by the smart marker
            Cell cell = _workbook.Worksheets[sheetIndex].Cells[rowIndex, colIndex];

            // Example modification logic:
            // - If the cell contains a numeric value, add 10.
            // - If the cell contains a string, append "_Modified".
            if (cell.Type == CellValueType.IsNumeric)
            {
                cell.PutValue(cell.DoubleValue + 10);
            }
            else if (cell.Type == CellValueType.IsString)
            {
                cell.PutValue(cell.StringValue + "_Modified");
            }
        }
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a new workbook (template) and add a smart marker
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            // Smart marker that will be replaced by data from Table1.Column1
            sheet.Cells["A1"].PutValue("&=Table1.Column1");

            // 2. Prepare a simple data source
            DataTable dt = new DataTable("Table1");
            dt.Columns.Add("Column1", typeof(string));
            dt.Rows.Add("Value1");
            dt.Rows.Add("Value2");

            // 3. Set up the WorkbookDesigner, assign the callback, and process smart markers
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook,
                CallBack = new MySmartMarkerCallback(workbook) // attach custom callback
            };
            designer.SetDataSource(dt);
            designer.Process(true); // process all smart markers

            // 4. Save the resulting workbook
            workbook.Save("SmartMarkerCallbackResult.xlsx");
        }
    }
}