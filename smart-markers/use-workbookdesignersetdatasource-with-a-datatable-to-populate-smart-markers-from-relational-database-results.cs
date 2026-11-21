// Title: Populate Excel smart markers from a DataTable with WorkbookDesigner.SetDataSource in Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells set DataTable as smart marker data source | Using WorkbookDesigner to fill an Excel template with database rows | Aspose.Cells example for binding relational DataTable to smart markers | How to process smart markers after assigning a DataTable in .NET | Populate smart markers in an .xlsx template from a DataTable
// Tags: WorkbookDesigner SetDataSource DataTable | smart markers populate relational data | Aspose.Cells bind DataTable to Excel template | process smart markers workbook | Excel template data binding Aspose.Cells

using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an Excel template containing smart markers, creates a WorkbookDesigner, builds a DataTable with employee records, assigns the DataTable as the smart‑marker data source via SetDataSource, processes the markers, and saves the populated workbook.
class SmartMarkerExample
{
    static void Main()
    {
        // Load the Excel template that contains smart markers
        Workbook workbook = new Workbook("SmartMarkerTemplate.xlsx"); // {LoadWorkbook}
        
        // Create a WorkbookDesigner for processing smart markers
        WorkbookDesigner designer = new WorkbookDesigner(workbook); // {CreateWorkbookDesigner}
        
        // Prepare a DataTable with relational data
        DataTable dt = new DataTable("Employees");
        dt.Columns.Add("ID", typeof(int));
        dt.Columns.Add("Name", typeof(string));
        dt.Columns.Add("Department", typeof(string));
        dt.Columns.Add("Salary", typeof(decimal));

        // Fill the DataTable with sample rows (replace with actual DB data)
        dt.Rows.Add(1, "John Doe", "Finance", 75000);
        dt.Rows.Add(2, "Jane Smith", "HR", 68000);
        dt.Rows.Add(3, "Bob Johnson", "IT", 82000);
        dt.Rows.Add(4, "Alice Brown", "Marketing", 71000);

        // Set the DataTable as the data source for smart markers
        designer.SetDataSource(dt); // {SetDataSource}
        
        // Process the smart markers in the workbook
        designer.Process(); // {ProcessSmartMarkers}
        
        // Save the populated workbook
        workbook.Save("SmartMarkerResult.xlsx"); // {SaveWorkbook}
    }
}
