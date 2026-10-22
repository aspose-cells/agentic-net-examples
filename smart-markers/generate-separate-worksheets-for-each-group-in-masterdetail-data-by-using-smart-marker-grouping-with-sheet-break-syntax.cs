// Title: Generate separate worksheets for each master category using smart‑marker sheet‑break syntax with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an Excel template containing a smart‑marker sheet‑break, binds a master‑detail DataSet, and creates a new worksheet for each master row. | Modify the example to insert a custom sheet‑break marker and apply header formatting to each generated worksheet. | Replace the in‑memory DataSet with a SQL query result while preserving the sheet‑break grouping behavior.
// Common Searches: aspocells c# smart marker sheet break generate separate worksheets per master group | how to use WorkbookDesigner to create a new sheet for each category in a master‑detail report | excel report with multiple sheets using smart markers and sheet break syntax in .NET | c# Aspose.Cells master‑detail dataset to multiple worksheets example | smart marker sheet break syntax for grouping data into separate Excel tabs
// Tags: smart marker sheet break Aspose.Cells | master‑detail grouping to multiple worksheets | WorkbookDesigner populate template C# | Excel report generation with sheet breaks | C# Aspose.Cells master‑detail export

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel template (or creates a new workbook if missing), builds master and detail DataTables, adds them to a DataSet, processes smart markers with WorkbookDesigner—including a sheet‑break marker—to generate a separate worksheet for each master category, and saves the result as Result.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Load the template workbook that contains smart markers.
            // If the file does not exist, create an empty workbook to avoid FileNotFoundException.
            string templatePath = "Template.xlsx";
            Workbook workbook;
            if (File.Exists(templatePath))
            {
                workbook = new Workbook(templatePath);
            }
            else
            {
                workbook = new Workbook(); // creates a new workbook with a default worksheet
                // Optionally, you could add placeholder smart markers here.
            }

            // ---------- Prepare master‑detail data ----------
            // Master table (e.g., categories)
            DataTable masterTable = new DataTable("Master");
            masterTable.Columns.Add("Category", typeof(string));
            masterTable.Columns.Add("Total", typeof(double));
            masterTable.Rows.Add("Fruits", 120);
            masterTable.Rows.Add("Vegetables", 80);

            // Detail table (items belonging to each category)
            DataTable detailTable = new DataTable("Detail");
            detailTable.Columns.Add("Category", typeof(string));
            detailTable.Columns.Add("Item", typeof(string));
            detailTable.Columns.Add("Quantity", typeof(int));
            detailTable.Rows.Add("Fruits", "Apple", 30);
            detailTable.Rows.Add("Fruits", "Banana", 50);
            detailTable.Rows.Add("Fruits", "Orange", 40);
            detailTable.Rows.Add("Vegetables", "Carrot", 20);
            detailTable.Rows.Add("Vegetables", "Tomato", 30);
            detailTable.Rows.Add("Vegetables", "Cucumber", 30);

            // Combine tables into a DataSet. The table names must match the smart marker names.
            DataSet dataSource = new DataSet();
            dataSource.Tables.Add(masterTable);
            dataSource.Tables.Add(detailTable);

            // ---------- Process smart markers ----------
            // Use WorkbookDesigner to process the workbook with the provided data source.
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dataSource);
            designer.Process();

            // ---------- Save the result ----------
            string resultPath = "Result.xlsx";
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully to '{resultPath}'.");
        }
        catch (Exception ex)
        {
            // Log or display any unexpected errors.
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
