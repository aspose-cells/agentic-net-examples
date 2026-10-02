// Title: Create hierarchical master‑detail tables in Excel using Aspose.Cells smart marker grouping with C#
// AI Prompts: Generate an Excel workbook that expands each master row and nests its related detail rows by applying smart marker grouping with a DataSet in Aspose.Cells C#. | Add extra detail columns to the smart marker template and reprocess the workbook so the hierarchical output includes the new fields. | Extend the template to introduce a second grouping level (e.g., product category) and produce a multi‑level grouped report using WorkbookDesigner.
// Common Searches: Aspose.Cells C# example for master‑detail smart marker grouping in Excel | How to use WorkbookDesigner with a DataSet to create hierarchical Excel reports | C# code to link master and detail tables via DataRelation for smart markers | Generate grouped rows in an Excel template using Aspose.Cells smart markers | Populate Excel template from multiple DataTables with a master‑detail relationship C#
// Tags: smart marker grouping with WorkbookDesigner | master-detail DataSet relation Aspose.Cells | hierarchical Excel generation C# | populate Excel template from DataTables | grouped rows using smart markers

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The sample loads a template workbook, builds master and detail DataTables, creates a DataRelation on OrderID, adds both tables to a DataSet, assigns the DataSet to a WorkbookDesigner, processes the smart markers to produce hierarchical rows, and saves the populated workbook as Result.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the template workbook
            string templatePath = "Template.xlsx";

            // Ensure the template file exists to avoid FileNotFoundException
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Error: Template file '{templatePath}' not found.");
                return;
            }

            // Load the Excel template that contains smart markers for grouping
            Workbook workbook = new Workbook(templatePath);

            // ---------- Prepare master data ----------
            DataTable masterTable = new DataTable("Master");
            masterTable.Columns.Add("OrderID", typeof(int));
            masterTable.Columns.Add("Customer", typeof(string));

            masterTable.Rows.Add(1, "Alice");
            masterTable.Rows.Add(2, "Bob");
            masterTable.Rows.Add(3, "Charlie");

            // ---------- Prepare detail data ----------
            DataTable detailTable = new DataTable("Detail");
            detailTable.Columns.Add("OrderID", typeof(int));
            detailTable.Columns.Add("Product", typeof(string));
            detailTable.Columns.Add("Quantity", typeof(int));

            detailTable.Rows.Add(1, "Apple", 10);
            detailTable.Rows.Add(1, "Banana", 5);
            detailTable.Rows.Add(2, "Orange", 7);
            detailTable.Rows.Add(2, "Grape", 3);
            detailTable.Rows.Add(3, "Mango", 12);
            detailTable.Rows.Add(3, "Peach", 8);

            // ---------- Combine into a DataSet and define relation ----------
            DataSet dataSet = new DataSet();
            dataSet.Tables.Add(masterTable);
            dataSet.Tables.Add(detailTable);

            // Relation on OrderID links master rows to their detail rows
            dataSet.Relations.Add(
                "MasterDetail",
                masterTable.Columns["OrderID"]!,
                detailTable.Columns["OrderID"]!
            );

            // ---------- Process smart markers ----------
            // Use WorkbookDesigner to process smart markers with the dataset
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dataSet);
            designer.Process();

            // Save the populated workbook
            string resultPath = "Result.xlsx";
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully to '{resultPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
