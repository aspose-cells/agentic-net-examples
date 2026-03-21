using System;
using System.Data;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the template XLSX that contains smart markers
        string templatePath = "TemplateWithSmartMarkers.xlsx";

        // Load the workbook (smart markers are recognized automatically)
        Workbook workbook = new Workbook(templatePath);

        // Initialize the WorkbookDesigner with the loaded workbook
        WorkbookDesigner designer = new WorkbookDesigner
        {
            Workbook = workbook
        };

        // Retrieve and display all smart markers found in the workbook
        string[] markers = designer.GetSmartMarkers();
        Console.WriteLine("Detected smart markers:");
        foreach (string marker in markers)
        {
            Console.WriteLine(marker);
        }

        // Prepare a data source that matches the smart markers in the template.
        DataTable dataTable = new DataTable("Data");
        dataTable.Columns.Add("Name", typeof(string));
        dataTable.Columns.Add("Age", typeof(int));
        dataTable.Rows.Add("Alice", 30);
        dataTable.Rows.Add("Bob", 25);

        // Set the data source for the designer
        designer.SetDataSource(dataTable);

        // Ensure the range that contains the smart markers is named "_CellsSmartMarkers"
        Worksheet sheet = workbook.Worksheets[0];
        Aspose.Cells.Range smartRange = sheet.Cells.CreateRange("A2:B3"); // adjust to actual marker area
        smartRange.Name = "_CellsSmartMarkers";

        // Process the smart markers and populate the data
        designer.Process();

        // Save the populated workbook
        string outputPath = "Result.xlsx";
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved to {outputPath}");
    }
}