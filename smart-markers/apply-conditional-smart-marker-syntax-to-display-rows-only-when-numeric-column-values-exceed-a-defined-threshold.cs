// Title: Use Aspose.Cells C# conditional smart markers to display rows only when a numeric column exceeds a threshold
// Common Searches: how to hide Excel rows with Aspose.Cells smart markers based on a numeric value in C# | Aspose.Cells example for filtering rows where Amount is greater than 500 using {if} smart marker | C# WorkbookDesigner {if} syntax to render rows only when a column meets a condition | using a DataTable as the data source for Aspose.Cells smart markers with conditional blocks | Aspose.Cells smart marker conditional block to show rows above a specified threshold
// Tags: WorkbookDesigner conditional rendering C# | if block smart marker Aspose.Cells | numeric threshold row filtering Excel | DataTable integration with Aspose.Cells smart markers | smart marker {if} syntax for Excel export

using Aspose.Cells;
using System.Data;

// The sample creates a workbook, adds Product and Amount headers, and inserts a smart marker row that uses the {if Amount > 500}{Product}{Amount}{endif} syntax. A DataTable with sample data is set as the data source for WorkbookDesigner, which processes the markers so only rows with Amount greater than 500 are rendered. The resulting file is saved as ConditionalSmartMarkerOutput.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Write column headers
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Amount");

        // Insert a smart marker row with a conditional block.
        // The row will be rendered only when the numeric column "Amount" exceeds the threshold (e.g., 500).
        // The {if ...} ... {endif} syntax controls the visibility.
        sheet.Cells["A2"].PutValue("{if Amount > 500}{Product}");
        sheet.Cells["B2"].PutValue("{Amount}{endif}");

        // Prepare a data source (DataTable) with sample data
        DataTable data = new DataTable();
        data.Columns.Add("Product", typeof(string));
        data.Columns.Add("Amount", typeof(double));

        data.Rows.Add("Apple", 300);   // Below threshold – row will be omitted
        data.Rows.Add("Banana", 600);  // Above threshold – row will be displayed
        data.Rows.Add("Cherry", 800);  // Above threshold – row will be displayed

        // Set the data source for the smart marker processor
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.SetDataSource(data);

        // Process the smart markers, applying the conditional logic
        designer.Process();

        // Save the resulting workbook
        workbook.Save("ConditionalSmartMarkerOutput.xlsx");
    }
}
