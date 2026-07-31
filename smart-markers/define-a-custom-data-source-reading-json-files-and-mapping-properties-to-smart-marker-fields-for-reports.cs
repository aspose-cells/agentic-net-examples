// Title: C# – Populate Excel Smart Markers from JSON with Aspose.Cells
// Description: Load an Excel template that contains smart markers, read a JSON file, bind the JSON string to the marker name using WorkbookDesigner.SetJsonDataSource, process the markers, and save the filled workbook as a report. The example demonstrates end‑to‑end JSON‑driven report generation in .NET.
// Keywords: Aspose.Cells | smart markers | JSON data source | C# .NET | WorkbookDesigner | SetJsonDataSource | Excel template automation | generate Excel report | populate workbook from JSON | global developers
// Common Searches: Aspose.Cells bind JSON to smart markers C# | WorkbookDesigner SetJsonDataSource example | fill Excel template with JSON data .NET | smart markers JSON mapping tutorial | generate report.xlsx from data.json using Aspose
// Developer Intent: Use Aspose.Cells to replace smart‑marker placeholders in an Excel template with values from a JSON file.
// Use Cases: Create a single report by merging a pre‑designed workbook with a JSON payload. | Batch‑produce multiple Excel reports by iterating over a collection of JSON files and a common template. | Expose an API endpoint that returns a dynamically generated Excel file based on incoming JSON data.
// AI Prompts: Generate C# code that reads a JSON file and applies WorkbookDesigner.SetJsonDataSource to fill smart markers named 'Data' in an Excel template. | Explain the steps to configure smart markers in Excel so they map to JSON properties when processed with Aspose.Cells. | List common reasons why JSON values might not appear after calling designer.Process() and how to troubleshoot them.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsJsonSmartMarkerDemo
{
    // Load an Excel template that contains smart markers, read a JSON file, bind the JSON string to the marker name using WorkbookDesigner.SetJsonDataSource, process the markers, and save the filled workbook as a report. The example demonstrates end‑to‑end JSON‑driven report generation in .NET.
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Load the Excel template that contains smart markers.
            //    Example smart markers in the template:
            //        A1 : &=Data.Name
            //        B1 : &=Data.Age
            //        C1 : &=Data.City
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook("Template.xlsx"); // template must exist in the executable folder

            // -----------------------------------------------------------------
            // 2. Create a WorkbookDesigner and associate it with the loaded workbook.
            // -----------------------------------------------------------------
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // -----------------------------------------------------------------
            // 3. Read JSON data from a file. The JSON represents an object with
            //    properties that correspond to the smart marker fields.
            // -----------------------------------------------------------------
            string jsonFilePath = "data.json";               // path to the JSON file
            string jsonContent = File.ReadAllText(jsonFilePath);

            // -----------------------------------------------------------------
            // 4. Set the JSON string as a data source for the smart markers.
            //    The first parameter is the data source name used in the markers.
            //    In the template we used the name "Data", therefore we pass it here.
            // -----------------------------------------------------------------
            designer.SetJsonDataSource("Data", jsonContent);

            // -----------------------------------------------------------------
            // 5. Process the smart markers – this populates the worksheet cells
            //    with values extracted from the JSON object.
            // -----------------------------------------------------------------
            designer.Process();

            // -----------------------------------------------------------------
            // 6. Save the populated workbook to a new file.
            // -----------------------------------------------------------------
            workbook.Save("Report.xlsx");
        }
    }
}
