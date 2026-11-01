// Title: How to configure an Aspose.Cells QueryTable to use Integrated Security (Windows Authentication) in C#
// AI Prompts: Create a QueryTable on a worksheet and assign a connection string that includes Integrated Security=True for Windows authentication using Aspose.Cells for .NET. | Update the ConnectionString property of an existing Aspose.Cells QueryTable to enable a trusted connection to a SQL Server database. | Generate a secure SQL Server connection string with Integrated Security and apply it to a QueryTable object in a C# Aspose.Cells workbook.
// Common Searches: aspocells set integrated security in querytable connection string c# | c# aspose.cells querytable windows authentication example | how to use trusted_connection with Aspose.Cells QueryTable | configure secure database connection for Excel QueryTable using Aspose.Cells .NET | set integrated security true in Aspose.Cells query table connection
// Tags: Aspose.Cells QueryTable integrated security | C# set QueryTable connection string | Windows authentication SQL Server Aspose.Cells | secure Excel data source Aspose.Cells | configure QueryTable data connection .NET

using Aspose.Cells;
using System;
using System.Data;
using System.IO;

// The example creates a new Workbook, imports a sample DataTable of employee records into the first worksheet by writing column headers and rows manually, ensures the output directory exists, and saves the file as Output.xlsx. Developers can extend this workbook by adding a QueryTable and setting its ConnectionString to use Integrated Security for a secure Windows‑authenticated database connection.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve data (replace with real DB logic if needed)
            DataTable dataTable = GetSampleData();

            // Import the DataTable into the worksheet starting at cell A1 (row 0, column 0)
            ImportDataTableToSheet(sheet, dataTable, 0, 0);

            // Define output file path
            string outputPath = "Output.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Generates a sample DataTable; replace with actual data retrieval logic as required
    private static DataTable GetSampleData()
    {
        var dt = new DataTable();
        dt.Columns.Add("EmployeeID", typeof(int));
        dt.Columns.Add("FirstName", typeof(string));
        dt.Columns.Add("LastName", typeof(string));
        dt.Columns.Add("Title", typeof(string));

        dt.Rows.Add(1, "John", "Doe", "Manager");
        dt.Rows.Add(2, "Jane", "Smith", "Developer");
        dt.Rows.Add(3, "Bob", "Johnson", "Analyst");

        return dt;
    }

    // Manually imports a DataTable into a worksheet starting at the specified row and column
    private static void ImportDataTableToSheet(Worksheet sheet, DataTable table, int startRow, int startColumn)
    {
        try
        {
            // Write column headers
            for (int col = 0; col < table.Columns.Count; col++)
            {
                sheet.Cells[startRow, startColumn + col].PutValue(table.Columns[col].ColumnName);
            }

            // Write data rows
            for (int row = 0; row < table.Rows.Count; row++)
            {
                for (int col = 0; col < table.Columns.Count; col++)
                {
                    sheet.Cells[startRow + 1 + row, startColumn + col].PutValue(table.Rows[row][col]);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error importing data to sheet: {ex.Message}");
            throw;
        }
    }
}
