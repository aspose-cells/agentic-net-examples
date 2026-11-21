// Title: Insert detail smart markers and process them with WorkbookDesigner in Aspose.Cells for .NET (C#)
// Common Searches: how to insert detail smart markers into an Excel template using Aspose.Cells C# | Aspose.Cells C# set DetailSheet parameter for smart marker processing on a specific worksheet | using WorkbookDesigner to bind a DataTable to smart markers in a .NET Excel workbook | C# example generating Excel output from smart markers with Aspose.Cells | process smart markers on a chosen worksheet with Aspose.Cells WorkbookDesigner
// Tags: insert detail smart markers Aspose.Cells | WorkbookDesigner smart marker processing C# | DetailSheet parameter Aspose.Cells | bind DataTable to smart markers .NET | generate Excel from template with smart markers

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The example loads a template workbook, adds «Orders.OrderID», «Orders.OrderDate», and «Orders.Amount» placeholders as detail smart markers, creates a DataTable with sample order data, assigns the table to a WorkbookDesigner, processes the smart markers, and saves the populated workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "template.xlsx";

            // Verify that the template file exists before attempting to load it
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file '{templatePath}' not found.");
                return;
            }

            // Load the workbook (template)
            Workbook workbook = new Workbook(templatePath);

            // Insert smart markers in the main sheet (assumed first sheet)
            Worksheet mainSheet = workbook.Worksheets[0];
            mainSheet.Cells["A1"].PutValue("«Orders.OrderID»");
            mainSheet.Cells["B1"].PutValue("«Orders.OrderDate»");
            mainSheet.Cells["C1"].PutValue("«Orders.Amount»");

            // Prepare a sample data source for the smart markers
            DataTable ordersTable = new DataTable("Orders");
            ordersTable.Columns.Add("OrderID", typeof(int));
            ordersTable.Columns.Add("OrderDate", typeof(DateTime));
            ordersTable.Columns.Add("Amount", typeof(double));
            ordersTable.Rows.Add(1001, DateTime.Now.AddDays(-2), 250.75);
            ordersTable.Rows.Add(1002, DateTime.Now.AddDays(-1), 180.00);
            ordersTable.Rows.Add(1003, DateTime.Now, 320.40);

            // Process the smart markers using WorkbookDesigner (replacement for SmartMarkerProcessor)
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(ordersTable);
            designer.Process();

            // Save the workbook
            const string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
