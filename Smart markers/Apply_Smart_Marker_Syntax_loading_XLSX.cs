using System;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add smart markers manually
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "DataSheet";

            // Place smart markers in cells
            sheet.Cells["A1"].PutValue("&Data.Name");
            sheet.Cells["B1"].PutValue("&Data.Age");

            // Initialize the WorkbookDesigner with the created workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Prepare a data source matching the smart marker names
            DataTable data = new DataTable("Data");
            data.Columns.Add("Name", typeof(string));
            data.Columns.Add("Age", typeof(int));
            data.Rows.Add("John Doe", 30);
            data.Rows.Add("Jane Smith", 28);

            // Bind the data source to the designer
            designer.SetDataSource(data);

            // Process the smart markers and populate the worksheet with data
            designer.Process();

            // Save the resulting workbook
            workbook.Save("SmartMarkerResult.xlsx");
        }
    }
}