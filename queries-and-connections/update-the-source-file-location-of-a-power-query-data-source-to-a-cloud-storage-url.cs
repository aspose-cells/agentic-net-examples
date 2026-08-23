// Title: Update a Power Query connection's source file to a cloud storage URL using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that scans a workbook’s DataConnections, finds Power Query connections, and sets their SourceFile to a given cloud URL. | Show how to log each updated Power Query connection name while changing its source path in an Excel file with Aspose.Cells. | Demonstrate saving the workbook after modifying Power Query connection URLs to a cloud location using Aspose.Cells. | Provide a reusable method that accepts a workbook path and a cloud URL, then updates all Power Query SourceFile properties.
// Common Searches: Aspose.Cells C# change Power Query connection source to cloud URL | programmatically update external data connection file path in Excel using Aspose.Cells | set SourceFile property for Power Query connections in .xlsx with C# | how to point Power Query source to an online Excel file using Aspose.Cells | iterate workbook.DataConnections and modify Power Query source location in .NET
// Tags: Power Query source URL Aspose.Cells | update external connection source C# | set SourceFile cloud storage | iterate DataConnections Power Query | modify Power Query file path .NET | Aspose.Cells change connection source

using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

// The example loads an existing workbook, loops through its DataConnections, identifies Power Query connections via the PowerQueryFormula property, updates each connection’s SourceFile to a specified cloud storage URL, logs the changes, and saves the workbook as a new file.
class UpdatePowerQuerySource
{
    static void Main()
    {
        // Load the workbook that contains Power Query connections
        Workbook workbook = new Workbook("input.xlsx");

        // Define the new cloud storage URL for the source file
        string newSourceUrl = "https://mycloudstorage.com/data/newsource.xlsx";

        // Iterate through all external data connections
        foreach (ExternalConnection connection in workbook.DataConnections)
        {
            // Identify Power Query connections by checking the PowerQueryFormula property
            if (connection.PowerQueryFormula != null)
            {
                // Update the SourceFile property to point to the cloud URL
                connection.SourceFile = newSourceUrl;
                Console.WriteLine($"Updated Power Query connection '{connection.Name}' to new source: {newSourceUrl}");
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
