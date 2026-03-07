using System;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerVariableArrayDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook to act as the template
            Workbook templateWorkbook = new Workbook();
            Worksheet sheet = templateWorkbook.Worksheets[0];

            // Add smart markers for the variable array (group name: COLORS_TIMES)
            // Header row (optional)
            sheet.Cells["A1"].PutValue("Color");
            sheet.Cells["B1"].PutValue("Times");

            // Smart marker rows – they will be expanded automatically
            sheet.Cells["A2"].PutValue("&=COLORS_TIMES.COLORS");
            sheet.Cells["B2"].PutValue("&=COLORS_TIMES.TIMES");

            // Create a WorkbookDesigner and attach the template workbook
            WorkbookDesigner designer = new WorkbookDesigner(templateWorkbook);

            // Prepare a DataTable that matches the smart marker group name (COLORS_TIMES)
            DataTable colorsTimesTable = new DataTable("COLORS_TIMES");
            colorsTimesTable.Columns.Add("COLORS", typeof(string));
            colorsTimesTable.Columns.Add("TIMES", typeof(int));

            // Fill the table with variable‑length data
            colorsTimesTable.Rows.Add("Red", 5);
            colorsTimesTable.Rows.Add("Green", 10);
            colorsTimesTable.Rows.Add("Blue", 15);
            colorsTimesTable.Rows.Add("Yellow", 20);

            // Set the DataTable as a data source for the designer
            designer.SetDataSource(colorsTimesTable);

            // Process the smart markers – the variable array will be expanded automatically
            designer.Process();

            // Optional: add another vertical array starting at cell D2 (row index 1, column index 3)
            string[] extraData = new string[] { "Extra1", "Extra2", "Extra3" };
            designer.Workbook.Worksheets[0].Cells.ImportArray(extraData, 1, 3, true);

            // Save the resulting workbook
            designer.Workbook.Save("ResultWithVariableArray.xlsx");
        }
    }
}