using Aspose.Cells;
using Aspose.Cells.ExternalConnections;
using System;

class UpdateDbConnectionDescription
{
    static void Main()
    {
        // Load the workbook that contains the DBConnection(s)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all external connections in the workbook
        foreach (ExternalConnection connection in workbook.DataConnections)
        {
            // Process only DBConnection objects
            if (connection is DBConnection dbConnection)
            {
                // Update the description to reflect the new purpose after migration
                dbConnection.ConnectionDescription = "Migrated connection for new data source";

                // Optional: output the change to the console
                Console.WriteLine($"Connection '{dbConnection.Name}' description updated.");
            }
        }

        // Save the workbook with the modified connection description
        workbook.Save("output.xlsx");
        Console.WriteLine("Workbook saved successfully.");
    }
}