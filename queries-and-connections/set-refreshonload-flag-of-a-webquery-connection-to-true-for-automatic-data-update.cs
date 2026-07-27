using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    public class SetRefreshOnLoadForWebQuery
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook();

                // Ensure there is at least one external connection.
                // In a real scenario you would add a WebQueryConnection to the collection.
                // Here we simply use the first connection if it exists.
                if (workbook.DataConnections.Count > 0 && workbook.DataConnections[0] is WebQueryConnection webQuery)
                {
                    // Set the RefreshOnLoad flag to true so the connection refreshes automatically
                    // when the workbook is opened.
                    webQuery.RefreshOnLoad = true;

                    Console.WriteLine("RefreshOnLoad set to: " + webQuery.RefreshOnLoad);
                }
                else
                {
                    Console.WriteLine("No external connections found in the workbook.");
                }

                // Save the workbook with the modified connection settings
                string outputPath = "WebQuery_RefreshOnLoad.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SetRefreshOnLoadForWebQuery.Run();
        }
    }
}