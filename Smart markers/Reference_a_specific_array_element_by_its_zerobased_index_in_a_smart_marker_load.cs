using System;
using System.Collections.Generic;
using Aspose.Cells;

class SmartMarkerArrayIndexDemo
{
    static void Main()
    {
        // Load the template workbook that contains smart markers.
        Workbook template = new Workbook("Template.xlsx");

        // Initialize WorkbookDesigner with the loaded workbook.
        WorkbookDesigner designer = new WorkbookDesigner();
        designer.Workbook = template;

        // Prepare a data source: a list of strings.
        List<string> names = new List<string> { "Alice", "Bob", "Charlie", "Diana" };

        // Set the data source for the smart marker. The name "Names" must match the marker prefix.
        designer.SetDataSource("Names", names);

        // Optional: display all smart markers found in the template.
        string[] markers = designer.GetSmartMarkers();
        Console.WriteLine("Smart markers found in the template:");
        foreach (string m in markers)
        {
            Console.WriteLine(m);
        }

        // Process the smart markers – this will replace "&=Names[2]" with "Charlie".
        designer.Process();

        // Save the resulting workbook.
        designer.Workbook.Save("Result.xlsx");
        Console.WriteLine("Processing complete. Result saved to Result.xlsx");
    }
}