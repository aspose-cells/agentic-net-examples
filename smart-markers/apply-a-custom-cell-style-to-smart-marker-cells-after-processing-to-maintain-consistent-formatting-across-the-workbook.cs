using System;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerStyleDemo
{
    class Program
    {
        static void Main()
        {
            // Load the template workbook that contains smart markers
            Workbook workbook = new Workbook("template.xlsx");

            // Initialize WorkbookDesigner with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };

            // ----- Prepare a sample data source -----
            DataTable dt = new DataTable("Employees");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Age", typeof(int));
            dt.Columns.Add("Department", typeof(string));

            dt.Rows.Add("John Doe", 30, "Sales");
            dt.Rows.Add("Jane Smith", 28, "Marketing");
            dt.Rows.Add("Bob Johnson", 45, "HR");

            // Set the data source for the smart markers
            designer.SetDataSource(dt);

            // Process all smart markers in the workbook
            designer.Process();

            // ----- Create a custom style to be applied after processing -----
            Style customStyle = workbook.CreateStyle();
            // Example: light yellow background with bold font
            customStyle.Pattern = BackgroundType.Solid;
            customStyle.ForegroundColor = System.Drawing.Color.LightYellow;
            customStyle.Font.IsBold = true;

            // Apply the custom style only to cells that already have a custom style
            // (i.e., cells that were populated by smart markers and may have inherited formatting)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                // Iterate through all used rows and columns
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];
                        // Check if the cell currently has a custom style
                        if (cell.HasCustomStyle)
                        {
                            // Apply the custom style while preserving explicitly set formatting
                            cell.SetStyle(customStyle, true);
                        }
                    }
                }
            }

            // Save the resulting workbook
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}