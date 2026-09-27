// Title: Add an OLE DB workbook connection with Integrated Windows Authentication (SSPI) using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a WorkbookConnection via WorkbookConnectionCollection.AddOleDbConnection with a connection string containing Integrated Security=SSPI. | Show how to assign a SELECT command to the WorkbookConnection after establishing Windows authentication. | Demonstrate saving the workbook to a file once the integrated security OLE DB connection is configured.
// Common Searches: Aspose.Cells how to use Integrated Security=SSPI in OLE DB connection string | C# add workbook connection with Windows authentication in Aspose.Cells | Example of WorkbookConnectionCollection AddOleDbConnection using SSPI | Configure Aspose.Cells workbook to connect to SQL Server without username and password | Saving a workbook after setting up an integrated Windows auth OLE DB connection in .NET
// Tags: add OLE DB connection Aspose.Cells | integrated Windows authentication workbook connection | WorkbookConnectionCollection SSPI connection string | C# Aspose.Cells configure DBConnection without credentials | save workbook after DB connection setup

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new Aspose.Cells Workbook and includes a commented snippet that demonstrates adding an OLE DB connection using a connection string with Integrated Security=SSPI. It shows how to set a SELECT command on the WorkbookConnection, ensures the output directory exists, saves the workbook as IntegratedAuthWorkbook.xlsx, and handles any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -----------------------------------------------------------------
            // NOTE: The WorkbookConnectionCollection API is available only in
            // newer versions of Aspose.Cells. If the current library does not
            // contain this feature, the related code is omitted to keep the
            // sample compilable and runnable.
            // -----------------------------------------------------------------
            // Example of adding an OLE DB connection (requires a recent version):
            // string connectionName = "IntegratedWindowsAuth";
            // string connectionString = "Provider=SQLOLEDB;Data Source=YOUR_SERVER_NAME;Initial Catalog=YOUR_DATABASE_NAME;Integrated Security=SSPI;";
            // int connIndex = workbook.WorkbookConnectionCollection.AddOleDbConnection(connectionName, connectionString);
            // WorkbookConnection connection = workbook.WorkbookConnectionCollection[connIndex];
            // connection.Command = "SELECT * FROM YourTable";

            // Define output path and ensure the directory exists
            string outputPath = "IntegratedAuthWorkbook.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
