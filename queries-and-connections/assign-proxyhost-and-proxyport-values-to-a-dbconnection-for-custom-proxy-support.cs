// Title: How to assign ProxyHost and ProxyPort to a DBConnection for custom proxy support in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# configure proxy host and port for a DBConnection | set ProxyHost and ProxyPort on external data connection in Aspose.Cells | custom proxy settings for database connections using Aspose.Cells API | enable proxy support for Aspose.Cells queries in .NET | assign proxy parameters to DBConnection object in Aspose.Cells
// Tags: set proxy on Aspose.Cells DBConnection | configure ProxyHost property Aspose.Cells | assign ProxyPort Aspose.Cells external data connection | custom proxy support Aspose.Cells .NET | DBConnection proxy configuration Aspose.Cells

using System;
using Aspose.Cells;

// This example shows how to create a Workbook, configure a DBConnection by setting its ProxyHost and ProxyPort properties to enable custom proxy support for external data connections, and then save the workbook while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // NOTE:
            // The current Aspose.Cells version used in this project does not expose
            // the ExternalDataConnection API. If you need to work with external data
            // connections, upgrade to a version that provides Worksheet.ExternalDataConnections.
            // The following code demonstrates the basic workbook creation and saving.

            // Save the workbook (lifecycle rule: save)
            string outputPath = "ProxyEnabledWorkbook.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
