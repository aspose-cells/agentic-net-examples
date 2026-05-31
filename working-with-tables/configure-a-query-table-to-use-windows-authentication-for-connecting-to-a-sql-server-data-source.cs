using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsQueryTableWindowsAuth
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string templatePath = "TemplateWithQueryTable.xlsx";
                const string outputPath = "OutputWithWindowsAuth.xlsx";

                // Verify that the template file exists before attempting to load it.
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file not found: {Path.GetFullPath(templatePath)}");
                    return;
                }

                // Load the workbook that already contains a query table.
                Workbook workbook = new Workbook(templatePath);

                // Access the first worksheet (adjust index if needed).
                Worksheet sheet = workbook.Worksheets[0];

                // Ensure the worksheet has at least one query table.
                if (sheet.QueryTables.Count == 0)
                {
                    Console.WriteLine("No query tables found in the worksheet.");
                    return;
                }

                // Get the first query table.
                QueryTable queryTable = sheet.QueryTables[0];

                // The query table is linked to an external DB connection.
                if (queryTable.ExternalConnection is DBConnection dbConnection)
                {
                    // Set the connection string for SQL Server (Windows Authentication).
                    dbConnection.ConnectionString =
                        "Data Source=YOUR_SERVER_NAME;Initial Catalog=YOUR_DATABASE_NAME;Integrated Security=SSPI;";

                    // Specify that the authentication method is Windows (Integrated).
                    dbConnection.CredentialsMethodType = CredentialsMethodType.Integrated;

                    Console.WriteLine("Query table connection configured to use Windows authentication.");
                }
                else
                {
                    Console.WriteLine("The external connection is not a DBConnection.");
                    return;
                }

                // Save the modified workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}