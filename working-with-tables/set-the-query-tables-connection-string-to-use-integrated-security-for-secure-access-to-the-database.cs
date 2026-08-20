// Title: Configure QueryTable Connection String for Integrated Security with Aspose.Cells (.NET)
// Description: Shows how to locate a QueryTable in a worksheet, set its ExternalConnection.ConnectionString to an OLE DB string that uses Windows Integrated Security (SSPI), and save the workbook.
// Keywords: Aspose.Cells | QueryTable | Integrated Security | Windows Authentication | OLE DB | ConnectionString | .NET | C# | Excel automation | secure database access | external connection
// Common Searches: Aspose.Cells set query table integrated security | How to use Windows authentication for QueryTable in Aspose.Cells | Change QueryTable connection string to SSPI C# | Update external connection of existing query table Aspose.Cells | Secure Excel data connection with Integrated Security Aspose.Cells
// Developer Intent: Update an existing QueryTable’s connection string to use Windows Integrated Security for secure data retrieval.
// Use Cases: Convert a QueryTable from SQL authentication to Windows authentication without storing credentials. | Programmatically generate Excel reports that pull data from a protected SQL Server using SSPI. | Automate workbook creation where the data connection must comply with corporate security policies.
// AI Prompts: Provide C# code using Aspose.Cells to modify a QueryTable’s ExternalConnection.ConnectionString to Integrated Security=SSPI. | Explain the steps to retrieve and change the connection string of an existing QueryTable in Aspose.Cells for Windows authentication. | Describe how to verify that a saved Excel file’s QueryTable uses Integrated Security when refreshed.

using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    // Shows how to locate a QueryTable in a worksheet, set its ExternalConnection.ConnectionString to an OLE DB string that uses Windows Integrated Security (SSPI), and save the workbook.
    public class SetQueryTableIntegratedSecurity
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add sample data (optional)
                sheet.Cells["A1"].PutValue("ID");
                sheet.Cells["B1"].PutValue("Name");
                sheet.Cells["A2"].PutValue(1);
                sheet.Cells["B2"].PutValue("Sample");

                // If a query table exists, update its connection string to use Integrated Security
                if (sheet.QueryTables.Count > 0)
                {
                    QueryTable queryTable = sheet.QueryTables[0];
                    ExternalConnection connection = queryTable.ExternalConnection;

                    // Example OLE DB connection string with Integrated Security
                    connection.ConnectionString =
                        "Provider=SQLOLEDB;Data Source=MyServer;Initial Catalog=MyDatabase;Integrated Security=SSPI;";
                }

                // Save the workbook
                string outputPath = "QueryTableIntegratedSecurity.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
