// Title: How to assign a default OLE DB connection string and SELECT command to a PivotTable’s DataSource with Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells in C# to set PivotTable.DataSource to a string array containing an OLE DB connection string and a SELECT statement for external workbook data. | Update an existing PivotTable so it pulls data from an external Excel file via a default connection string without recreating the workbook. | Programmatically configure a PivotTable to retrieve data from an external data source by assigning a connection string and SQL query through Aspose.Cells.
// Common Searches: aspnet set pivot table data source to external workbook using aspose.cells | c# assign ole db connection string to pivot table data source asp.net | how to use Aspose.Cells to link a pivot table to an external Excel file | configure PivotTable.DataSource with connection string and SQL query in C# | example of external data source for Aspose.Cells pivot table
// Tags: Aspose.Cells PivotTable.DataSource connection string | C# external OLE DB data source for PivotTable | set pivot table source to external Excel file Aspose | assign string array to PivotTable.DataSource Aspose.Cells | configure external data retrieval for Aspose.Cells pivot

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample loads or creates an Excel workbook, ensures a PivotTable exists, defines an OLE DB connection string and a SELECT command for an external workbook, assigns them to the PivotTable's DataSource as a string array, and saves the modified file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            Workbook workbook;

            // Load existing workbook if present; otherwise create a sample workbook
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = CreateSampleWorkbook();
            }

            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure at least one pivot table exists
            if (worksheet.PivotTables.Count == 0)
            {
                AddSamplePivotTable(worksheet);
            }

            PivotTable pivotTable = worksheet.PivotTables[0];

            // Example external data source connection string (adjust as needed)
            string externalConnection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=ExternalData.xlsx;Persist Security Info=False;";

            // External command (e.g., select all data from the first sheet)
            string externalCommand = "SELECT * FROM [Sheet1$]";

            // Assign external data source to the pivot table (requires string array)
            pivotTable.DataSource = new string[] { externalConnection, externalCommand };

            // Save the modified workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Creates a workbook with sample data and a pivot table
    private static Workbook CreateSampleWorkbook()
    {
        Workbook wb = new Workbook();
        Worksheet dataSheet = wb.Worksheets[0];
        dataSheet.Name = "Data";

        // Sample data
        dataSheet.Cells["A1"].PutValue("Category");
        dataSheet.Cells["B1"].PutValue("Amount");
        dataSheet.Cells["A2"].PutValue("A");
        dataSheet.Cells["B2"].PutValue(100);
        dataSheet.Cells["A3"].PutValue("B");
        dataSheet.Cells["B3"].PutValue(200);
        dataSheet.Cells["A4"].PutValue("A");
        dataSheet.Cells["B4"].PutValue(150);
        dataSheet.Cells["A5"].PutValue("B");
        dataSheet.Cells["B5"].PutValue(250);

        // Add a pivot table on a new sheet
        int pivotSheetIdx = wb.Worksheets.Add();
        Worksheet pivotSheet = wb.Worksheets[pivotSheetIdx];
        pivotSheet.Name = "Pivot";

        string sourceData = dataSheet.Name + "!A1:B5";
        string destCell = "A3";

        int ptIndex = pivotSheet.PivotTables.Add(sourceData, destCell, "PivotTable1");
        PivotTable pt = pivotSheet.PivotTables[ptIndex];

        // Note: Adding fields is optional for this example; omitted to avoid API version issues
        return wb;
    }

    // Adds a simple pivot table to an existing worksheet if none exists
    private static void AddSamplePivotTable(Worksheet dataSheet)
    {
        Workbook wb = dataSheet.Workbook;
        int pivotSheetIdx = wb.Worksheets.Add();
        Worksheet pivotSheet = wb.Worksheets[pivotSheetIdx];
        pivotSheet.Name = "Pivot";

        string sourceData = dataSheet.Name + "!A1:B5";
        string destCell = "A3";

        int ptIndex = pivotSheet.PivotTables.Add(sourceData, destCell, "PivotTable1");
        PivotTable pt = pivotSheet.PivotTables[ptIndex];

        // Note: Adding fields is optional for this example; omitted to avoid API version issues
    }
}
