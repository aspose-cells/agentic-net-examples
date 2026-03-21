using System;
using System.Data;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the template workbook that contains a date smart marker (e.g., "&=Data.Date")
        Workbook workbook = new Workbook("template.xlsx");

        // Access the worksheet and the cell with the smart marker
        Worksheet sheet = workbook.Worksheets[0];
        Cell smartMarkerCell = sheet.Cells["A2"]; // adjust the address as needed

        // Apply the desired date format (MM/dd/yyyy) to the cell
        Style dateStyle = smartMarkerCell.GetStyle();
        dateStyle.Custom = "MM/dd/yyyy";
        smartMarkerCell.SetStyle(dateStyle);

        // Prepare a data source with a DateTime column
        DataTable data = new DataTable("Data");
        data.Columns.Add("Date", typeof(DateTime));
        data.Rows.Add(DateTime.Now);
        data.Rows.Add(DateTime.Now.AddDays(1));
        data.Rows.Add(DateTime.Now.AddDays(2));

        // Set up the WorkbookDesigner, bind the data source, and process smart markers
        WorkbookDesigner designer = new WorkbookDesigner
        {
            Workbook = workbook
        };
        designer.SetDataSource("Data", data);
        designer.Process();

        // Save the resulting workbook
        workbook.Save("output.xlsx");
    }
}