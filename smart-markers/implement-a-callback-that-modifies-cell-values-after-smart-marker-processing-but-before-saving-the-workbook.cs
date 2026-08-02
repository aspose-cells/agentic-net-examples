using System;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerCallbackDemo
{
    // Custom callback that modifies cell values after each smart marker is processed
    public class MySmartMarkerCallback : ISmartMarkerCallBack
    {
        private readonly Workbook _workbook;

        // Receive the workbook instance so we can access its cells
        public MySmartMarkerCallback(Workbook workbook)
        {
            _workbook = workbook;
        }

        // This method is called for every smart marker cell during processing
        public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
        {
            // Get the cell that was just populated by the smart marker
            Cell cell = _workbook.Worksheets[sheetIndex].Cells[rowIndex, colIndex];

            // Example modification logic:
            // - If the cell contains a string, prepend "Modified-"
            // - If the cell contains a numeric value, multiply it by 10
            // - Otherwise leave it unchanged
            if (cell.Type == CellValueType.IsString)
            {
                string original = cell.StringValue;
                cell.PutValue("Modified-" + original);
            }
            else if (cell.Type == CellValueType.IsNumeric)
            {
                double original = cell.DoubleValue;
                cell.PutValue(original * 10);
            }
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Load a workbook that contains smart markers (e.g., "&=Employees.Name")
            Workbook workbook = new Workbook("SmartMarkerTemplate.xlsx");

            // Create a WorkbookDesigner and assign the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };

            // Prepare sample data source
            DataTable dt = new DataTable("Employees");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Salary", typeof(double));
            dt.Rows.Add("John Doe", 1200.5);
            dt.Rows.Add("Jane Smith", 1500.0);

            // Set the data source for the designer
            designer.SetDataSource(dt);

            // Assign the custom callback to modify cells after smart marker processing
            designer.CallBack = new MySmartMarkerCallback(workbook);

            // Process the smart markers (true = preserve unrecognized markers)
            designer.Process(true);

            // Save the workbook after the callback has modified the cells
            workbook.Save("SmartMarkerProcessed.xlsx");
        }
    }
}