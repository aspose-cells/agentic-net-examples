// Title: C# – Update Power Query connection source to a cloud URL using Aspose.Cells
// Description: Loads an Excel workbook, finds Power Query connections via the PowerQueryFormula property, sets each connection's SourceFile to a specified cloud storage URL, logs the change, and saves the workbook.
// Keywords: Aspose.Cells | Power Query | external data connection | C# | cloud storage URL | SourceFile property | update connection path | programmatic Excel | .NET
// Common Searches: change Power Query source to web URL Aspose.Cells C# | update external connection file path in Excel programmatically | set Power Query SourceFile to cloud storage using .NET | Aspose.Cells modify Power Query connection URL | batch update Power Query data source in multiple workbooks
// Developer Intent: Replace the local file path of Power Query connections with a cloud storage URL programmatically.
// Use Cases: Shift on‑premises Excel reports to reference data stored in a cloud bucket without manual edits. | Automate the re‑pointing of dozens of workbooks to a new shared cloud file after a data migration. | Maintain a single source of truth for Power Query data by directing all templates to a centralized cloud location.
// AI Prompts: Generate C# code that uses Aspose.Cells to set the SourceFile of every Power Query connection in a workbook to a given cloud URL and saves the file. | Explain how to confirm that each Power Query connection was successfully updated to the cloud URL with Aspose.Cells APIs. | Provide error‑handling patterns for missing PowerQueryFormula properties or invalid URLs when updating connections.

using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

// Loads an Excel workbook, finds Power Query connections via the PowerQueryFormula property, sets each connection's SourceFile to a specified cloud storage URL, logs the change, and saves the workbook.
class UpdatePowerQuerySource
{
    static void Main()
    {
        // Load the workbook that contains the Power Query connection
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all external data connections in the workbook
        foreach (ExternalConnection connection in workbook.DataConnections)
        {
            // Identify Power Query connections by checking the PowerQueryFormula property
            if (connection.PowerQueryFormula != null)
            {
                // Update the source file location to the desired cloud storage URL
                connection.SourceFile = "https://mycloudstorage.com/data/sourcefile.xlsx";
                Console.WriteLine($"Updated SourceFile for connection '{connection.Name}' to cloud URL.");
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
