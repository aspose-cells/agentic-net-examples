using System;
using System.Data;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Header
        cells["A1"].PutValue("Date");

        // Smart marker with date format "MMMM dd, yyyy"
        // The format after the colon is applied to the date value during processing
        cells["A2"].PutValue("&=[ReportDate:MMMM dd, yyyy]");

        // Prepare a data source containing DateTime values
        DataTable data = new DataTable("Report");
        data.Columns.Add("ReportDate", typeof(DateTime));
        data.Rows.Add(new DateTime(2023, 5, 15));
        data.Rows.Add(new DateTime(2023, 12, 1));

        // Set the data source for the smart markers and process them
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.SetDataSource(data);
        designer.Process();

        // Save the resulting workbook
        workbook.Save("SmartMarkerDateFormat.xlsx");
    }
}