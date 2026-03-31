using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the XLSX workbook that may contain smart markers
        string filePath = "template.xlsx";

        // Load the workbook from the specified file
        Workbook workbook = new Workbook(filePath);

        // Initialize WorkbookDesigner with the loaded workbook
        WorkbookDesigner designer = new WorkbookDesigner();
        designer.Workbook = workbook;

        // Retrieve all smart markers defined in the workbook
        string[] smartMarkers = designer.GetSmartMarkers();

        // Output the smart markers to the console
        Console.WriteLine("Smart markers found in the workbook:");
        foreach (string marker in smartMarkers)
        {
            Console.WriteLine(marker);
        }
    }
}