using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class WebQueryExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add header cells
        sheet.Cells["A1"].PutValue("Column1");
        sheet.Cells["B1"].PutValue("Column2");

        // Create and configure a WebQueryConnection
        WebQueryConnection webConn = new WebQueryConnection(workbook, "SampleWebQuery", "https://example.com/data.html");
        webConn.IsHtmlTables = true;
        webConn.IsXml = false;
        webConn.RefreshOnLoad = true;
        webConn.SaveData = true;

        // Add the connection to the workbook
        workbook.DataConnections.Add(webConn);

        // Add a QueryTable that uses the created WebQueryConnection
        QueryTable queryTable = sheet.QueryTables.Add("SampleWebQueryTable", webConn.Name, 1, 0);

        // Save the workbook
        workbook.Save("WebQueryDemo.xlsx");
    }
}