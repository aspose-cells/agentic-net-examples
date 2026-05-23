using System;
using System.Data;
using System.IO;
using Aspose.Cells;

public class WorkbookDesignerDataTableDemo
{
    public static void Main()
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create a DataTable and populate it with sample data
        DataTable dataTable = new DataTable("Products");
        dataTable.Columns.Add("ProductID", typeof(int));
        dataTable.Columns.Add("ProductName", typeof(string));
        dataTable.Columns.Add("Price", typeof(decimal));

        dataTable.Rows.Add(1, "Laptop", 1200.00m);
        dataTable.Rows.Add(2, "Smartphone", 800.00m);
        dataTable.Rows.Add(3, "Tablet", 450.00m);

        // Initialize WorkbookDesigner with a new workbook
        WorkbookDesigner designer = new WorkbookDesigner
        {
            Workbook = new Workbook()
        };

        // Add smart markers to the first worksheet
        Worksheet sheet = designer.Workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("&=$ProductID");
        sheet.Cells["B1"].PutValue("&=$ProductName");
        sheet.Cells["C1"].PutValue("&=$Price");

        // Assign the DataTable as the data source for the designer
        designer.SetDataSource(dataTable);

        // Process the smart markers and populate the worksheet with data
        designer.Process();

        // Define output file path
        string outputPath = "DataTableOutput.xlsx";

        // Ensure the directory for the output file exists
        string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
        if (!Directory.Exists(outputDir))
        {
            Directory.CreateDirectory(outputDir);
        }

        // Save the resulting workbook
        designer.Workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved to {outputPath}");
    }
}