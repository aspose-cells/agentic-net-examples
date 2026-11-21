// Title: Bind a DataTable to WorkbookDesigner and generate Excel with smart markers using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set DataTable as data source for WorkbookDesigner smart markers | example of using WorkbookDesigner.SetDataSource with a DataTable in .NET | how to generate an Excel file from a DataTable using smart markers in Aspose.Cells | load an existing Excel template or create a new workbook before processing smart markers in C#
// Tags: WorkbookDesigner SetDataSource with DataTable | Aspose.Cells smart markers custom data source | C# generate Excel from DataTable using smart markers | load or create workbook template Aspose.Cells | process smart markers after binding data

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The example creates a DataTable of employee records, loads an existing Excel template or creates a new workbook, assigns the DataTable to WorkbookDesigner via SetDataSource, processes smart markers to merge the data, and saves the resulting workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a DataTable and populate it with sample data
            DataTable dt = new DataTable("Employees");
            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Salary", typeof(double));

            dt.Rows.Add(1, "John Doe", 50000);
            dt.Rows.Add(2, "Jane Smith", 60000);
            dt.Rows.Add(3, "Bob Johnson", 55000);

            // Load an existing workbook template if it exists; otherwise create a new workbook
            Workbook workbook;
            string templatePath = "Template.xlsx";
            if (File.Exists(templatePath))
            {
                workbook = new Workbook(templatePath);
            }
            else
            {
                workbook = new Workbook(); // creates an empty workbook
            }

            // Initialize WorkbookDesigner with the workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Assign the DataTable as the data source for the designer
            designer.SetDataSource(dt);

            // Process the designer to merge the data into the workbook
            designer.Process();

            // Save the resulting workbook
            string outputPath = "Output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
