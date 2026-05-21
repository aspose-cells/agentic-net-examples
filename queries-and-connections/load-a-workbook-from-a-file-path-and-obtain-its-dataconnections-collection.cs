using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class LoadWorkbookAndGetDataConnections
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Path to the Excel file to be loaded
            string filePath = "input.xlsx";

            // Verify that the file exists before loading
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                // Load the workbook from the specified file path
                Workbook workbook = new Workbook(filePath);

                // Obtain the DataConnections collection from the loaded workbook
                var dataConnections = workbook.DataConnections;

                // Display the number of data connections present in the workbook
                Console.WriteLine("DataConnections count: " + dataConnections.Count);

                // Iterate through each connection (if any) and display its name
                for (int i = 0; i < dataConnections.Count; i++)
                {
                    var connection = dataConnections[i];
                    Console.WriteLine($"Connection {i + 1}: {connection.Name}");
                }

                // Example of saving the workbook after inspection (optional)
                // workbook.Save("output.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}