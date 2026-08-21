// Title: Aspose.Cells C# – Enable Background Refresh for External Data Connections
// Description: Demonstrates how to set the BackgroundRefresh property to true on an external data connection in an Aspose.Cells workbook, allowing asynchronous refreshes that keep the UI responsive, then saves the file.
// Keywords: Aspose.Cells background refresh | C# external data connection async | Enable BackgroundRefresh property | Aspose.Cells workbook save | asynchronous data refresh .NET
// Common Searches: Aspose.Cells enable background refresh C# | set BackgroundRefresh true external connection | asynchronous refresh of external data in Aspose.Cells | how to keep UI responsive while refreshing workbook data
// Developer Intent: Configure an external data connection to refresh asynchronously using Aspose.Cells for .NET.
// Use Cases: Improve UI responsiveness by refreshing data in the background. | Apply background refresh only when a workbook contains external connections. | Persist the asynchronous refresh setting by saving the workbook after configuration.
// AI Prompts: Generate C# code that iterates over all external data connections in an Aspose.Cells workbook, sets BackgroundRefresh = true for each, and initiates an asynchronous refresh. | Provide an example that handles the case where a workbook has no external connections while attempting to enable BackgroundRefresh, including proper error messages.

using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    // Demonstrates how to set the BackgroundRefresh property to true on an external data connection in an Aspose.Cells workbook, allowing asynchronous refreshes that keep the UI responsive, then saves the file.
    public class EnableBackgroundRefreshDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Check if the workbook contains any external data connections
                if (workbook.DataConnections.Count > 0)
                {
                    // Access the first external connection
                    ExternalConnection connection = workbook.DataConnections[0];

                    // Enable background refresh so the connection can be refreshed asynchronously
                    connection.BackgroundRefresh = true;

                    Console.WriteLine("BackgroundRefresh set to: " + connection.BackgroundRefresh);
                }
                else
                {
                    // No external connections exist; inform the user.
                    Console.WriteLine("The workbook does not contain any external data connections.");
                }

                // Save the workbook (lifecycle rule: save)
                string outputPath = "EnableBackgroundRefreshDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            EnableBackgroundRefreshDemo.Run();
        }
    }
}
