// Title: Use Aspose.Cells C# smart markers with {if} condition to show rows only when a numeric column exceeds a threshold
// AI Prompts: Write C# code that creates an Excel workbook with Aspose.Cells, inserts a smart‑marker row using the {if} syntax to display the row only when the Amount column is greater than a specified value, and binds a DataTable as the data source. | Show how to configure WorkbookDesigner in C# to apply conditional smart markers for numeric comparisons, including setting up the marker tags and processing the workbook.
// Common Searches: aspocells c# conditional smart marker row based on amount value | how to hide Excel rows with Aspose.Cells smart markers when a column is below a threshold | using {if} smart marker syntax in Aspose.Cells to filter rows by numeric column | c# example of WorkbookDesigner processing smart markers with numeric condition
// Tags: Aspose.Cells conditional smart marker | WorkbookDesigner conditional row | Excel row visibility numeric threshold | smart marker numeric comparison C#

using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example creates a workbook, adds header cells, and places a smart‑marker row that uses the {if Amount > 500}{endif} syntax so rows appear only when the Amount exceeds 500. A DataTable supplies sample data, WorkbookDesigner binds the data source, processes the conditional markers, and the result is saved as an Excel file.
class ConditionalSmartMarkerExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Set up header row
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Amount");

        // Insert smart marker row with conditional syntax.
        // The row will be displayed only when the Amount value exceeds the threshold (e.g., 500).
        // {if Amount > 500} starts the condition, {endif} ends it.
        // Place the start tag in the first cell and the end tag in the last cell of the same row.
        sheet.Cells["A2"].PutValue("{if Amount > 500}{Product}");
        sheet.Cells["B2"].PutValue("{Amount}{endif}");

        // Prepare data source
        DataTable dt = new DataTable("Data");
        dt.Columns.Add("Product", typeof(string));
        dt.Columns.Add("Amount", typeof(double));

        // Sample data
        dt.Rows.Add("Apple", 300);   // Below threshold – row will be hidden
        dt.Rows.Add("Banana", 750);  // Above threshold – row will be shown
        dt.Rows.Add("Cherry", 1200); // Above threshold – row will be shown
        dt.Rows.Add("Date", 150);    // Below threshold – row will be hidden

        // Apply the data source to the workbook using WorkbookDesigner
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.SetDataSource(dt);
        designer.Process(); // Process smart markers

        // Save the result
        workbook.Save("ConditionalSmartMarkerOutput.xlsx");
    }
}
