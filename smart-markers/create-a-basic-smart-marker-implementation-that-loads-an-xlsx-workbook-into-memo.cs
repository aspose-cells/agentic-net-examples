using System;
using System.Data;
using System.IO;
using Aspose.Cells;

class SmartMarkerExample
{
    static void Main()
    {
        // Path to the template workbook that contains smart markers
        string templatePath = "template.xlsx";

        // Load the template file into a memory stream (in‑memory processing)
        byte[] fileBytes = File.ReadAllBytes(templatePath);
        using (MemoryStream memoryStream = new MemoryStream(fileBytes))
        {
            // Load the workbook from the memory stream
            Workbook workbook = new Workbook(memoryStream);

            // Access the first worksheet (where smart markers are placed)
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range that contains the smart markers.
            // The range must be named "_CellsSmartMarkers" for range‑based processing.
            Aspose.Cells.Range smartRange = sheet.Cells.CreateRange("A2:B2");
            smartRange.Name = "_CellsSmartMarkers";

            // Create a WorkbookDesigner and assign the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = workbook;

            // Prepare a simple data source (DataTable) for the smart markers
            DataTable data = new DataTable("MyData");
            data.Columns.Add("Name", typeof(string));
            data.Columns.Add("Value", typeof(double));
            data.Rows.Add("Item1", 123.45);
            data.Rows.Add("Item2", 678.90);

            // Set the data source for the designer
            designer.SetDataSource(data);

            // Process only the defined smart‑marker range (true = preserve unrecognized markers)
            designer.Process(smartRange, true);

            // Save the processed workbook to a file
            workbook.Save("output.xlsx");
        }
    }
}