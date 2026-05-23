using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Markup;

namespace SmartMarkerStyleDemo
{
    // Callback that applies a predefined style to each cell populated by a smart marker
    public class StyleApplyingCallback : ISmartMarkerCallBack
    {
        private readonly Workbook _workbook;
        private readonly Style _style;

        public StyleApplyingCallback(Workbook workbook, Style style)
        {
            _workbook = workbook;
            _style = style;
        }

        // This method is invoked for every smart‑marker cell during processing
        public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
        {
            // Retrieve the target cell and apply the custom style
            Cell cell = _workbook.Worksheets[sheetIndex].Cells[rowIndex, colIndex];
            cell.SetStyle(_style);
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Load the template workbook that contains smart markers
            Workbook workbook = new Workbook("template.xlsx");

            // -----------------------------------------------------------------
            // Prepare a simple data source (replace with your actual source)
            // -----------------------------------------------------------------
            DataTable dt = new DataTable("Employees");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Age", typeof(int));
            dt.Rows.Add("John Doe", 30);
            dt.Rows.Add("Jane Smith", 28);

            // -----------------------------------------------------------------
            // Create the custom style that will be applied after processing
            // -----------------------------------------------------------------
            Style customStyle = workbook.CreateStyle();
            customStyle.Pattern = BackgroundType.Solid;
            customStyle.ForegroundColor = System.Drawing.Color.LightYellow;
            customStyle.Font.IsBold = true;
            customStyle.Font.Color = System.Drawing.Color.DarkBlue;

            // -----------------------------------------------------------------
            // Set up the WorkbookDesigner, data source, and the callback
            // -----------------------------------------------------------------
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook,
                CallBack = new StyleApplyingCallback(workbook, customStyle)
            };
            designer.SetDataSource(dt);

            // Process all smart markers; the callback will style each populated cell
            designer.Process();

            // Save the result workbook
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}