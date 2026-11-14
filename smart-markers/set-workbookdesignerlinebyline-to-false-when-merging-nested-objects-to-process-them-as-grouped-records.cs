// Title: Disable WorkbookDesigner.LineByLine to process nested parent‑child tables as grouped records with Aspose.Cells for .NET
// Common Searches: Aspose.Cells set WorkbookDesigner.LineByLine false for parent child dataset | How to merge nested DataSet tables using smart markers in Aspose.Cells C# | Process Excel template with grouped records instead of line‑by‑line mode in Aspose.Cells | Disable line‑by‑line processing for nested objects in WorkbookDesigner | Aspose.Cells smart markers example with parent‑child relationship
// Tags: WorkbookDesigner linebyline false | smart markers grouped records | parent child dataset Aspose.Cells | csharp excel template processing | aspose.cells nested objects merging

using System;
using System.Data;
using Aspose.Cells;

// The example loads an Excel template, creates a WorkbookDesigner, sets LineByLine to false so nested parent‑child tables are treated as grouped records, builds a DataSet with related Parent and Child tables, assigns it as the data source, processes the smart markers, and saves the resulting workbook.
class Program
{
    static void Main()
    {
        // Load the template workbook (replace with your template path)
        Workbook workbook = new Workbook("Template.xlsx");

        // Initialize WorkbookDesigner with the loaded workbook
        WorkbookDesigner designer = new WorkbookDesigner(workbook);

        // Set LineByLine to false so nested objects are processed as grouped records
        designer.LineByLine = false;

        // ----- Prepare a DataSet with nested (parent‑child) tables -----
        DataSet dataSet = new DataSet();

        // Parent table
        DataTable parentTable = new DataTable("Parent");
        parentTable.Columns.Add("Id", typeof(int));
        parentTable.Columns.Add("Name", typeof(string));
        parentTable.Rows.Add(1, "Alice");
        parentTable.Rows.Add(2, "Bob");
        dataSet.Tables.Add(parentTable);

        // Child table (related to Parent)
        DataTable childTable = new DataTable("Child");
        childTable.Columns.Add("ParentId", typeof(int));
        childTable.Columns.Add("Order", typeof(string));
        childTable.Rows.Add(1, "Order1");
        childTable.Rows.Add(1, "Order2");
        childTable.Rows.Add(2, "Order3");
        dataSet.Tables.Add(childTable);

        // Define the relation between Parent and Child tables
        dataSet.Relations.Add("Parent_Child", parentTable.Columns["Id"], childTable.Columns["ParentId"]);

        // Set the DataSet as the data source for the designer
        designer.SetDataSource(dataSet);

        // Process the template with the provided data
        designer.Process();

        // Save the resulting workbook (replace with your desired output path)
        workbook.Save("Result.xlsx");
    }
}
