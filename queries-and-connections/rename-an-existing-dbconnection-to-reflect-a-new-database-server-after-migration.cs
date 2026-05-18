using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    public class RenameDbConnectionDemo
    {
        public static void Run()
        {
            // Load the workbook that contains the existing DBConnection
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all external connections in the workbook
            foreach (ExternalConnection connection in workbook.DataConnections)
            {
                // Process only DBConnection objects
                if (connection is DBConnection dbConn)
                {
                    // Example: replace old server identifier "OldServer" with "NewServer" in the connection name
                    if (!string.IsNullOrEmpty(dbConn.Name) && dbConn.Name.Contains("OldServer"))
                    {
                        dbConn.Name = dbConn.Name.Replace("OldServer", "NewServer");
                    }

                    // Optionally, also update the connection string to point to the new server
                    if (!string.IsNullOrEmpty(dbConn.ConnectionString) && dbConn.ConnectionString.Contains("OldServer"))
                    {
                        dbConn.ConnectionString = dbConn.ConnectionString.Replace("OldServer", "NewServer");
                    }
                }
            }

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            RenameDbConnectionDemo.Run();
        }
    }
}