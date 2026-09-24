// Title: Generate an Excel report with Aspose.Cells C# by mapping a SQL Server view to smart‑marker tags
// AI Prompts: Write C# code that reads data from a SQL Server view into a DataTable, assigns it to a smart‑marker source called 'Employees' via WorkbookDesigner, places &={Employees}.Id, &={Employees}.Name, &={Employees}.Salary tags in worksheet cells, processes the markers, and saves the workbook as an XLSX file. | Show how to stream rows from a SqlDataReader directly into Aspose.Cells smart markers without loading the entire result set into memory, using WorkbookDesigner.SetDataSource with an enumerable data source. | Add robust error handling and logging to the smart‑marker example while using a real connection string and SqlDataAdapter to populate the DataTable from a view, keeping the marker mapping unchanged.
// Common Searches: Aspose.Cells C# example for binding a SQL Server view to smart markers and exporting to Excel | How to fill a DataTable from a database view and use it with WorkbookDesigner smart markers | Streaming rows from SqlDataReader directly into Aspose.Cells smart markers without a full DataTable | Smart marker syntax for Id, Name, Salary columns in an Excel template using Aspose.Cells | Setting up a custom data source for smart markers in Aspose.Cells C#
// Tags: Aspose.Cells WorkbookDesigner set data source from SQL view | C# smart markers map DataTable columns to Excel cells | Export XLSX using smart markers and database view | Stream SqlDataReader rows to Aspose.Cells smart markers | Excel report generation with smart markers in .NET

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The example demonstrates how to define a logical data source name, create a DataTable that matches smart‑marker columns, insert &={Employees}.Id, &={Employees}.Name, and &={Employees}.Salary tags into a worksheet, register the DataTable with WorkbookDesigner as a smart‑marker source, process the markers, and save the resulting workbook as an XLSX file.
public class SmartMarkerExample
{
    public static void Main()
    {
        // Logical name used inside the smart markers
        const string dataSourceName = "Employees";

        try
        {
            // Retrieve data (sample data in this example)
            DataTable dataTable = GetSampleDataTable();

            // Create a new workbook (no external template required)
            var workbook = new Workbook();

            // Access the first worksheet and set its name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Report";

            // Insert smart marker tags that map to data source columns
            // Example: assume the data source has columns Id, Name, Salary
            sheet.Cells["A1"].PutValue($"&={dataSourceName}.Id");
            sheet.Cells["B1"].PutValue($"&={dataSourceName}.Name");
            sheet.Cells["C1"].PutValue($"&={dataSourceName}.Salary");

            // Process smart markers using WorkbookDesigner
            var designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dataSourceName, dataTable);
            designer.Process();

            // Save the resulting workbook
            const string outputPath = "SmartMarkerReport.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Report saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Helper method to create a sample DataTable (replaces DB access)
    private static DataTable GetSampleDataTable()
    {
        var table = new DataTable();

        try
        {
            // Define columns matching the smart marker tags
            table.Columns.Add("Id", typeof(int));
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Salary", typeof(decimal));

            // Add sample rows
            table.Rows.Add(1, "Alice", 75000m);
            table.Rows.Add(2, "Bob", 62000m);
            table.Rows.Add(3, "Charlie", 54000m);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Data preparation error: {ex.Message}");
            throw;
        }

        return table;
    }
}
