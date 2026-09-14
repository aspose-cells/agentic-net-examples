// Title: Enable asynchronous background refresh for an external data connection in an Aspose.Cells workbook (C#)
// AI Prompts: Write C# code that opens a workbook, iterates over its DataConnections, sets BackgroundRefresh = true for each, and saves the file using Aspose.Cells. | Generate a C# snippet that checks whether a workbook contains any external connections, enables background refresh on the first connection, and logs the operation outcome. | Create a C# example that toggles the BackgroundRefresh property of an ExternalConnection object and persists the changes to an .xlsx workbook.
// Common Searches: Aspose.Cells C# enable background refresh for workbook data connections | how to set BackgroundRefresh on external connection using Aspose.Cells .NET | asynchronous data refresh in Excel file with Aspose.Cells API | C# sample to save workbook after modifying external connection settings in Aspose.Cells
// Tags: Aspose.Cells background refresh | Aspose.Cells external data connection | C# set BackgroundRefresh property | asynchronous data refresh Aspose.Cells | save workbook after connection update

using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

// This example demonstrates how to use Aspose.Cells for .NET to detect external data connections in a workbook, enable the BackgroundRefresh flag to allow asynchronous updates, and then save the modified workbook as an .xlsx file while providing console logging and error handling.
class EnableBackgroundRefreshDemo
{
    public static void Run()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Check if the workbook contains any external connections
            if (workbook.DataConnections.Count > 0)
            {
                // Access the first external connection
                ExternalConnection connection = workbook.DataConnections[0];

                // Enable background refresh (asynchronous update)
                connection.BackgroundRefresh = true;

                Console.WriteLine("BackgroundRefresh set to: " + connection.BackgroundRefresh);
            }
            else
            {
                Console.WriteLine("No external connections found in the workbook.");
            }

            // Save the workbook with the modified connection settings
            workbook.Save("BackgroundRefreshEnabled.xlsx");
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}

class Program
{
    static void Main(string[] args)
    {
        EnableBackgroundRefreshDemo.Run();
    }
}
