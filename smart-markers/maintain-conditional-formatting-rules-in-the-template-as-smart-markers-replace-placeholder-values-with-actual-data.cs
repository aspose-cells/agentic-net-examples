// Title: Keep Conditional Formatting When Processing Smart Markers with Aspose.Cells for .NET
// Description: Shows how to load a template workbook that already contains smart markers and conditional‑formatting rules, bind a DataTable to the "Employees" smart marker, process the markers with WorkbookDesigner while retaining the formatting, verify the rule count before and after processing, and save the final file.
// Keywords: Aspose.Cells | C# | .NET | WorkbookDesigner | smart markers | conditional formatting | preserve formatting | data binding | template workbook
// Common Searches: Aspose.Cells keep conditional formatting after smart marker processing | WorkbookDesigner preserve formatting C# | verify conditional formatting count before and after processing | smart markers data binding without losing formatting | how to retain conditional rules in Aspose.Cells template
// Developer Intent: Retain existing conditional‑formatting rules while populating a workbook using smart markers and a data source.
// Use Cases: Load a pre‑designed template that includes conditional formatting and smart markers, then fill it with data without altering the formatting. | Log the number of conditional‑formatting collections before and after WorkbookDesigner.Process to confirm they remain unchanged. | Apply the preserve‑unrecognized‑markers option to keep any markers that are not bound while still applying conditional formatting.
// AI Prompts: Write C# code that processes smart markers in an Aspose.Cells workbook and ensures all conditional formatting rules stay intact. | Explain how to check the conditional‑formatting count in a worksheet before and after calling WorkbookDesigner.Process. | Suggest a strategy for handling multiple worksheets, each with its own conditional formatting, when using smart markers in Aspose.Cells.

using System;
using System.Data;
using Aspose.Cells;

// Shows how to load a template workbook that already contains smart markers and conditional‑formatting rules, bind a DataTable to the "Employees" smart marker, process the markers with WorkbookDesigner while retaining the formatting, verify the rule count before and after processing, and save the final file.
class ConditionalFormattingSmartMarkerDemo
{
    static void Main()
    {
        // Load the template workbook that already contains smart markers and conditional formatting rules
        Workbook workbook = new Workbook("Template.xlsx");

        // Optional: check how many conditional formatting collections exist before processing
        Worksheet sheet = workbook.Worksheets[0];
        Console.WriteLine("Conditional formatting count before processing: " + sheet.ConditionalFormattings.Count);

        // Prepare a data source that matches the smart marker names used in the template
        DataTable dt = new DataTable("Employees");
        dt.Columns.Add("Name", typeof(string));
        dt.Columns.Add("Age", typeof(int));
        dt.Columns.Add("Department", typeof(string));
        dt.Rows.Add("John Doe", 30, "Sales");
        dt.Rows.Add("Jane Smith", 28, "HR");

        // Set up the WorkbookDesigner, bind the data source, and process the smart markers
        WorkbookDesigner designer = new WorkbookDesigner
        {
            Workbook = workbook
        };
        designer.SetDataSource("Employees", dt);
        designer.Process(true); // true = preserve unrecognized markers (set as needed)

        // Verify that conditional formatting rules are still present after processing
        Console.WriteLine("Conditional formatting count after processing: " + sheet.ConditionalFormattings.Count);

        // Save the processed workbook
        workbook.Save("Result.xlsx");
    }
}
