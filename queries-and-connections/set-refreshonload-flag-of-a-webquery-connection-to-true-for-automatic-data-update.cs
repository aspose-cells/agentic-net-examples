using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    public class SetRefreshOnLoadForWebQueryDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (in-memory)
                Workbook workbook = new Workbook();

                // Ensure there is at least one external connection of type WebQueryConnection
                if (workbook.DataConnections.Count > 0 && workbook.DataConnections[0] is WebQueryConnection webQuery)
                {
                    // Set the RefreshOnLoad flag to true so the connection refreshes automatically on opening
                    webQuery.RefreshOnLoad = true;
                    Console.WriteLine("RefreshOnLoad has been set to: " + webQuery.RefreshOnLoad);
                }
                else
                {
                    Console.WriteLine("No WebQueryConnection found in the workbook.");
                }

                // Save the workbook with the modified connection settings
                string outputPath = "WebQuery_RefreshOnLoad_True.xlsx";

                // If a file with the same name exists, it will be overwritten
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
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
            SetRefreshOnLoadForWebQueryDemo.Run();
        }
    }
}