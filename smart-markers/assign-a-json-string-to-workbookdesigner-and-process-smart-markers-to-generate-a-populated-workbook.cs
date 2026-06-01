using System;
using Aspose.Cells;

namespace AsposeCellsJsonSmartMarkerDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (empty template)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a smart marker that references a JSON field named "Name"
            // The marker syntax "&=$DataSource.Name" tells the designer to replace it with the value of the "Name" property
            sheet.Cells["A1"].PutValue("&=$DataSource.Name");

            // Insert another smart marker for "Age"
            sheet.Cells["A2"].PutValue("&=$DataSource.Age");

            // Create a WorkbookDesigner and assign the workbook to it
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = workbook;

            // Sample JSON data source
            string json = "{\"Name\":\"John Doe\",\"Age\":30,\"City\":\"New York\"}";

            // Set the JSON string as a data source named "DataSource"
            // The first parameter is the name used in the smart markers (DataSource)
            designer.SetJsonDataSource("DataSource", json);

            // Process the smart markers – this populates the cells with data from the JSON
            designer.Process();

            // Save the populated workbook to a file
            workbook.Save("JsonSmartMarkerOutput.xlsx");
        }
    }
}