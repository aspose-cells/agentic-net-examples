using System;
using Aspose.Cells;

class SmartMarkersDemo
{
    static void Main()
    {
        // Load an existing XLSX file that contains smart markers
        string inputPath = "template.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Initialize WorkbookDesigner with the loaded workbook
        WorkbookDesigner designer = new WorkbookDesigner();
        designer.Workbook = workbook;

        // Retrieve all smart markers defined in the workbook
        string[] smartMarkers = designer.GetSmartMarkers();

        // Display the smart markers
        Console.WriteLine("Smart markers found in the workbook:");
        foreach (string marker in smartMarkers)
        {
            Console.WriteLine(marker);
        }

        // Save the workbook (optional, no processing performed here)
        string outputPath = "output.xlsx";
        designer.Workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved to {outputPath}");
    }
}