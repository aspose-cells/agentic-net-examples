// Title: How to bind a custom DataTable to WorkbookDesigner for smart‑marker processing in Aspose.Cells for .NET
// AI Prompts: Write C# code that creates an in‑memory DataTable, loads an Excel template, assigns the DataTable to WorkbookDesigner via SetDataSource, processes smart markers, and saves the resulting workbook. | Explain the steps required to replace WorkbookDesigner’s default data source with a custom DataTable when generating Excel reports using Aspose.Cells. | Show how to populate a DataTable with employee information and use it to fill smart‑marker placeholders in a template workbook through WorkbookDesigner.
// Common Searches: aspnet bind datatable to workbookdesigner setdatasource aspocells | c# use workbookdesigner to fill smart markers from a datatable | how to generate excel from template using aspocells and custom datatable | set custom data source for workbookdesigner smart markers c# example | aspocells workbookdesigner process template with in‑memory datatable
// Tags: WorkbookDesigner SetDataSource with DataTable | Aspose.Cells smart markers from DataTable | C# generate Excel from template using WorkbookDesigner | in‑memory DataTable as data source for Aspose.Cells | process Excel template with custom data source

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Demonstrates creating a DataTable, loading a template workbook, assigning the DataTable as the data source to WorkbookDesigner, processing smart markers, and saving the resulting workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // 1. Create a custom DataTable and populate it with sample data
            DataTable dt = new DataTable("Employees");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Age", typeof(int));
            dt.Columns.Add("Department", typeof(string));

            dt.Rows.Add("Alice", 30, "HR");
            dt.Rows.Add("Bob", 45, "IT");
            dt.Rows.Add("Charlie", 28, "Finance");

            // 2. Initialize WorkbookDesigner
            WorkbookDesigner designer = new WorkbookDesigner();

            // 3. Load the template workbook (replace with your actual template path)
            designer.Workbook = new Workbook("Template.xlsx");

            // 4. Set the custom DataTable as the data source for the designer
            designer.SetDataSource(dt);

            // 5. Process the template to merge data
            designer.Process();

            // 6. Save the resulting workbook (replace with your desired output path)
            designer.Workbook.Save("Result.xlsx");
        }
    }
}
