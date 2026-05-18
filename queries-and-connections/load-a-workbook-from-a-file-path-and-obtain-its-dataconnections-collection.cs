using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class LoadWorkbookAndGetDataConnections
    {
        public static void Run()
        {
            // Path to the existing Excel file
            string filePath = "example.xlsx";

            // Load the workbook from the specified file path
            Workbook workbook = new Workbook(filePath);

            // Obtain the collection of external data connections in the workbook
            var dataConnections = workbook.DataConnections;

            // Display the number of data connections
            Console.WriteLine("DataConnections count: " + dataConnections.Count);

            // Iterate through the connections and display their names
            for (int i = 0; i < dataConnections.Count; i++)
            {
                var connection = dataConnections[i];
                Console.WriteLine($"Connection {i + 1}: {connection.Name}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            LoadWorkbookAndGetDataConnections.Run();
        }
    }
}