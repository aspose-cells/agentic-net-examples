// Title: Generate row‑specific totals with a formula parameter in Aspose.Cells smart markers (C#)
// AI Prompts: Write C# code that loads an Excel template, creates a DataTable containing Quantity, Price, and a Total column defined by a formula, and configures WorkbookDesigner to apply a formula parameter so each generated row computes Total correctly. | Explain how to set up a smart marker in an Excel template that uses a formula parameter to automatically adjust cell references when processing a DataSet with multiple rows in Aspose.Cells. | Show the steps to save the processed workbook to a new file after the smart marker formulas have been applied, including error handling for missing template files.
// Common Searches: how to use formula parameters with Aspose.Cells smart markers in C# | Aspose.Cells WorkbookDesigner adjust formulas for each row from DataSet | C# example of dynamic total calculation using smart markers and formula parameters
// Tags: WorkbookDesigner dynamic formula generation | smart marker row‑wise formula adjustment | C# create Excel from template with formulas | Aspose.Cells adjust cell references per row | DataSet driven total calculation smart marker

using System;
using System.IO;
using System.Data;
using Aspose.Cells;

// The example loads an Excel template containing smart markers, builds a DataSet with a DataTable that includes a formula column for total calculations, assigns the data source to a WorkbookDesigner, processes the smart markers so the formula references are automatically updated for each inserted row, and saves the resulting workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the template workbook that contains the smart markers
            string templatePath = "SmartMarkerTemplate.xlsx";

            // Verify that the template file exists to avoid FileNotFoundException
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Error: Template file '{templatePath}' not found.");
                return;
            }

            // Load the template workbook
            Workbook workbook = new Workbook(templatePath);

            // Prepare the data source as a DataSet with a DataTable named "Items"
            DataTable itemsTable = new DataTable("Items");
            itemsTable.Columns.Add("Product", typeof(string));
            itemsTable.Columns.Add("Quantity", typeof(int));
            itemsTable.Columns.Add("Price", typeof(double));
            itemsTable.Columns.Add("Total", typeof(string)); // formula string

            // Add rows (the formula references will be adjusted by Aspose.Cells)
            itemsTable.Rows.Add("Pen", 10, 1.5, "=C2*D2");
            itemsTable.Rows.Add("Notebook", 5, 3.0, "=C3*D3");
            itemsTable.Rows.Add("Eraser", 20, 0.5, "=C4*D4");

            DataSet dataSet = new DataSet();
            dataSet.Tables.Add(itemsTable);

            // Initialize the WorkbookDesigner and set the data source
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dataSet);

            // Apply smart markers (process the template)
            designer.Process();

            // Save the resulting workbook
            string resultPath = "SmartMarkerResult.xlsx";
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully to '{resultPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
