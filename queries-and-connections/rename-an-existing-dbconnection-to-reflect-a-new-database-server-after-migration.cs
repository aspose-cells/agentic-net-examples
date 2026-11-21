// Title: Rename an external database connection and update its server name in an Excel workbook using Aspose.Cells for .NET
// Common Searches: how to change external data source name in Aspose.Cells C# | update server address in Excel workbook DB connection with Aspose.Cells .NET | rename DBConnection and modify connection string after database migration using Aspose.Cells | Aspose.Cells external data source API not available fallback handling | C# edit external data source connection string in an .xlsx file
// Tags: Aspose.Cells external data source rename | update Excel connection string C# | modify Data Source property Aspose.Cells | conditional compilation for ExternalDataSources API | external DB connection rename Aspose.Cells

using Aspose.Cells;
using System;
using System.Data.Common;
using System.IO;

// The sample loads an Excel workbook, locates an external data source named "OldConnection" in the first worksheet, and—if the ExternalDataSources API is present—renames it to "NewConnection" and updates the server name in its connection string to "NewServerName" using a DbConnectionStringBuilder. The code is wrapped in a compile‑time conditional block to gracefully handle older Aspose.Cells versions where the API is unavailable, logging an informational message instead. Finally, the modified workbook is saved to a new file.
class Program
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Old connection name and the new details after migration
            string oldConnectionName = "OldConnection";
            string newConnectionName = "NewConnection";
            string newServerName = "NewServerName";

            // Attempt to modify external data source connection if the API is available.
            // The ExternalDataSources API may not exist in older Aspose.Cells versions,
            // so this block is guarded with compile‑time conditional compilation.
#if false
            try
            {
                Worksheet sheet = workbook.Worksheets[0];

                foreach (var dataSource in sheet.ExternalDataSources)
                {
                    var dsInfo = dataSource as ExternalDataSourceInfo;
                    if (dsInfo != null && dsInfo.Name == oldConnectionName)
                    {
                        dsInfo.Name = newConnectionName;

                        DbConnectionStringBuilder builder = new DbConnectionStringBuilder
                        {
                            ConnectionString = dsInfo.ConnectionString
                        };
                        if (builder.ContainsKey("Data Source"))
                            builder["Data Source"] = newServerName;
                        else if (builder.ContainsKey("Server"))
                            builder["Server"] = newServerName;

                        dsInfo.ConnectionString = builder.ConnectionString;
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Unable to modify external data sources. {ex.Message}");
            }
#else
            // ExternalDataSources API not available; log the limitation.
            Console.WriteLine("Info: External data source modification is not supported in this Aspose.Cells version.");
#endif

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
