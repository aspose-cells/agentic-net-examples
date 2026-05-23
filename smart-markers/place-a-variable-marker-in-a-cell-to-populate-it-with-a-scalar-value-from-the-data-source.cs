using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Place a smart marker in cell A1 that will be replaced by the data source value
        // The marker syntax &=$Name tells Aspose.Cells to look for a field named "Name"
        worksheet.Cells["A1"].PutValue("&=$Name");

        // Prepare a simple data source with a scalar value
        var dataSource = new { Name = "John Doe" };

        // Initialize WorkbookDesigner with the workbook
        WorkbookDesigner designer = new WorkbookDesigner(workbook);

        // Associate the data source with the name "Data" (required for smart markers)
        designer.SetDataSource("Data", dataSource);

        // Process the smart markers and populate the cell with the scalar value
        designer.Process();

        // Save the resulting workbook
        workbook.Save("VariableMarkerOutput.xlsx");
    }
}