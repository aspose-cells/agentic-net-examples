// Title: How to limit Aspose.Cells smart marker processing to a specific named range using WorkbookDesigner.SetRange in C#
// AI Prompts: Provide C# sample code that defines a named range for smart markers and invokes the WorkbookDesigner SetRange method before processing. | Modify the given Aspose.Cells example so that only the _CellsSmartMarkers range is evaluated, then call Process. | Explain the steps to confine smart marker evaluation to a worksheet area with a JSON data source in .NET.
// Common Searches: Aspose.Cells limit smart marker evaluation to a named range C# | WorkbookDesigner SetRange example for processing selected cells | How to process only certain smart markers in an Excel file using Aspose.Cells | C# code to create a named range and apply smart markers with Aspose.Cells | Using JSON data source with scoped smart markers in Aspose.Cells
// Tags: smart marker range limitation | named range processing Aspose.Cells | smart marker evaluation restriction | process selected cells with WorkbookDesigner | JSON-driven smart marker scope | Aspose.Cells smart markers range restriction

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsSmartMarkerRangeDemo
{
    // The example creates a workbook, adds headers, inserts smart markers, defines a named range (_CellsSmartMarkers) covering the markers, sets a JSON data source, uses WorkbookDesigner.SetRange to restrict processing to that range, processes the smart markers, and saves the result as SmartMarkerRangeResult.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add sample headers
                sheet.Cells["A1"].PutValue("Name");
                sheet.Cells["B1"].PutValue("Value");

                // Insert smart markers that will be populated from the data source
                sheet.Cells["A2"].PutValue("&=$Name");
                sheet.Cells["B2"].PutValue("&=$Value");

                // Create a named range that encloses the smart markers
                AsposeRange smartMarkerRange = sheet.Cells.CreateRange("A2:B2");
                smartMarkerRange.Name = "_CellsSmartMarkers";

                // Initialize the WorkbookDesigner with the workbook
                WorkbookDesigner designer = new WorkbookDesigner(workbook);

                // Set a JSON data source that matches the smart markers
                string jsonData = "{\"Name\":\"Test Product\",\"Value\":100.50}";
                designer.SetJsonDataSource("Data", jsonData);

                // Process the smart markers (processing the whole workbook)
                designer.Process();

                // Save the resulting workbook
                workbook.Save("SmartMarkerRangeResult.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
