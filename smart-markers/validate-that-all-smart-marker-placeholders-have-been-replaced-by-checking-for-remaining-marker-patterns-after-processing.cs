// Title: C# – Verify All Smart Markers Are Replaced Using WorkbookDesigner.GetSmartMarkers in Aspose.Cells
// Description: Shows how to load an Excel template containing smart markers, bind a DataTable as the data source, run WorkbookDesigner.Process, retrieve any leftover tokens with GetSmartMarkers, report the outcome, and save the completed workbook.
// Keywords: Aspose.Cells | smart markers | WorkbookDesigner | GetSmartMarkers | .NET | C# example | validate marker replacement | detect unreplaced placeholders | Excel template processing | automated reporting
// Common Searches: Aspose.Cells GetSmartMarkers example | check smart marker replacement C# | validate smart markers after Process | find remaining smart markers in Excel | C# Aspose.Cells smart marker validation
// Developer Intent: Ensure that no smart‑marker placeholders remain after processing the workbook.
// Use Cases: Load a pre‑designed template, bind data, process markers, then confirm replacement before publishing the file. | Add validation to a nightly report generator that throws an exception if any markers are left unreplaced. | Log unreplaced smart markers to a diagnostics file to aid debugging of dynamic document creation.
// AI Prompts: Generate C# code that processes smart markers with WorkbookDesigner and raises an InvalidOperationException when GetSmartMarkers returns any items. | Explain the behavior of GetSmartMarkers after calling Process and how to interpret its string array result. | Provide a snippet that writes each remaining smart marker to a log file instead of the console in an Aspose.Cells workflow.

using System;
using System.Data;
using Aspose.Cells;

// Shows how to load an Excel template containing smart markers, bind a DataTable as the data source, run WorkbookDesigner.Process, retrieve any leftover tokens with GetSmartMarkers, report the outcome, and save the completed workbook.
class SmartMarkerValidation
{
    static void Main()
    {
        // Load the template workbook that contains smart markers
        Workbook workbook = new Workbook("template.xlsx");

        // Initialize the WorkbookDesigner with the loaded workbook
        WorkbookDesigner designer = new WorkbookDesigner();
        designer.Workbook = workbook;

        // Prepare a data source that matches the smart markers in the template
        DataTable dt = new DataTable("Employees");
        dt.Columns.Add("Name", typeof(string));
        dt.Columns.Add("Age", typeof(int));
        dt.Rows.Add("John Doe", 30);
        dt.Rows.Add("Jane Smith", 28);

        // Bind the data source to the designer
        designer.SetDataSource(dt);

        // Process all smart markers in the workbook
        designer.Process();

        // Retrieve any remaining smart markers after processing
        string[] remainingMarkers = designer.GetSmartMarkers();

        // Validate that no placeholders remain
        if (remainingMarkers.Length == 0)
        {
            Console.WriteLine("All smart markers have been successfully replaced.");
        }
        else
        {
            Console.WriteLine("Unreplaced smart markers found:");
            foreach (string marker in remainingMarkers)
            {
                Console.WriteLine(marker);
            }
            // Optionally, you could throw an exception here
            // throw new InvalidOperationException("Smart marker replacement incomplete.");
        }

        // Save the processed workbook
        workbook.Save("output.xlsx");
    }
}
