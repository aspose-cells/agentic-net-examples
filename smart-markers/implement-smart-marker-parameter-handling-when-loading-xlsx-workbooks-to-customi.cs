using System;
using System.Data;
using System.Drawing;
using Aspose.Cells;

namespace SmartMarkerCallbackDemo
{
    class Program
    {
        static void Main()
        {
            // Load the template workbook that contains smart markers
            Workbook workbook = new Workbook("TemplateWithSmartMarkers.xlsx");

            // Initialize the WorkbookDesigner with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook,
                // Assign a custom callback to control how each smart marker value is inserted
                CallBack = new CustomSmartMarkerCallback(workbook)
            };

            // Prepare a data source (DataTable) that matches the smart markers in the template
            DataTable dt = new DataTable("Products");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Price", typeof(double));
            dt.Columns.Add("Quantity", typeof(int));
            dt.Rows.Add("Laptop", 1200.5, 3);
            dt.Rows.Add("Phone", 699.99, 5);

            // Bind the data source to the designer
            designer.SetDataSource(dt);

            // Process the smart markers; true = preserve unrecognized markers
            designer.Process(true);

            // Save the processed workbook
            designer.Workbook.Save("ProcessedOutput.xlsx");
        }
    }

    // Custom implementation of ISmartMarkerCallBack to modify insertion behavior
    class CustomSmartMarkerCallback : ISmartMarkerCallBack
    {
        private readonly Workbook _workbook;

        public CustomSmartMarkerCallback(Workbook workbook)
        {
            _workbook = workbook;
        }

        public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
        {
            // Access the cell that Aspose.Cells has just populated
            Worksheet sheet = _workbook.Worksheets[sheetIndex];
            Cell cell = sheet.Cells[rowIndex, colIndex];
            object originalValue = cell.Value;

            // Example: format the "Price" column as currency
            if (columnName.Equals("Price", StringComparison.OrdinalIgnoreCase) && originalValue is double price)
            {
                cell.PutValue(string.Format("${0:N2}", price));
            }
            // Example: highlight low stock in the "Quantity" column
            else if (columnName.Equals("Quantity", StringComparison.OrdinalIgnoreCase) && originalValue is int qty)
            {
                if (qty < 5)
                {
                    Style style = cell.GetStyle();
                    style.ForegroundColor = Color.LightCoral;
                    style.Pattern = BackgroundType.Solid;
                    cell.SetStyle(style);
                }
            }

            // Additional custom logic can be added here based on tableName or other criteria
        }
    }
}