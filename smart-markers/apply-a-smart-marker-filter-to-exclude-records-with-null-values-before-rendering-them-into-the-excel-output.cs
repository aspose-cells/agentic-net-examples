// Title: Use WorkbookDesigner smart‑marker filter to skip rows with DBNull values in an Excel template (C# Aspose.Cells)
// AI Prompts: Generate C# code that loads an Excel template, assigns a DataTable containing nullable fields to WorkbookDesigner, and applies a smart‑marker filter so rows with DBNull are omitted before saving. | Explain the smart‑marker syntax required to filter out null values and show how WorkbookDesigner processes the filter during workbook generation.
// Common Searches: asp.net c# workbookdesigner filter out rows with null values using smart markers | aspocells smart marker filter syntax to exclude DBNull rows | how to prevent empty rows when populating Excel template with DataTable and smart markers | skip rows with missing data in Aspose.Cells smart markers | using &? filter in Aspose.Cells smart markers for null handling
// Tags: WorkbookDesigner smart marker null filter | Aspose.Cells exclude DBNull rows | C# smart marker conditional row omission | Excel template data source filtering Aspose.Cells | smart marker filter syntax for null values

using System;
using System.Data;
using Aspose.Cells;

// Loads a template workbook, creates a DataTable with a DBNull entry, sets it as the data source for WorkbookDesigner, uses a smart‑marker filter (e.g., &?Value?) to omit rows where the Value column is null, processes the markers, and saves the filtered result.
class Program
{
    static void Main()
    {
        // Load the Excel template that contains smart markers with a filter (e.g., &Value?)
        Workbook workbook = new Workbook("Template.xlsx");

        // Create a DataTable that includes some null values
        DataTable table = new DataTable();
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Value", typeof(double));

        table.Rows.Add("Item1", 10);
        table.Rows.Add("Item2", DBNull.Value); // This row should be excluded by the filter
        table.Rows.Add("Item3", 30);

        // Assign the DataTable as the data source for smart markers
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.SetDataSource(table);

        // Process the smart markers; rows where the filtered field is null will be omitted
        designer.Process();

        // Save the final workbook
        workbook.Save("Result.xlsx");
    }
}
