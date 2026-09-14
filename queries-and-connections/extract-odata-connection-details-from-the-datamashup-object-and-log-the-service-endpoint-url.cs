// Title: How to extract OData service endpoint URLs from Power Query connections in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that opens an .xlsx file, reads its DataMashup, and prints each Power Query name together with its OData service URL. | Modify the example to return a Dictionary<string,string> where the key is the Power Query name and the value is the extracted endpoint URL, handling cases where the Url property is missing. | Create a reusable static method GetPowerQueryEndpoints(string workbookPath) that loads the workbook and returns a list of (query name, endpoint URL) tuples using reflection on the PowerQueryFormula objects.
// Common Searches: aspnet extract OData endpoint from Power Query DataMashup using Aspose.Cells | c# list all Power Query connection URLs in an Excel file with Aspose.Cells | how to read DataMashup PowerQueryFormulas Url property in Aspose.Cells .NET | retrieve OData service URLs from Excel workbook Power Query using reflection
// Tags: Aspose.Cells extract OData endpoint from DataMashup | C# read PowerQueryFormulas Url property | list Power Query connections in Excel workbook | handle missing Url property reflection | dictionary of query names to service URLs Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.QueryTables; // For DataMashup related types

// Loads an Excel workbook, accesses its DataMashup, iterates over PowerQueryFormulas, and logs each query name with the OData service endpoint URL obtained via reflection, handling absent Url properties gracefully.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook that may contain Power Query (OData) connections
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the DataMashup object; it can be null if no Power Query data exists
            DataMashup mashup = workbook.DataMashup;
            if (mashup == null)
            {
                Console.WriteLine("No Power Query (DataMashup) information found in the workbook.");
                return;
            }

            // Ensure the collection of PowerQueryFormulas is available
            var formulas = mashup.PowerQueryFormulas;
            if (formulas == null || formulas.Count == 0)
            {
                Console.WriteLine("The workbook does not contain any Power Query formulas.");
                return;
            }

            // Iterate through all Power Query formulas in the mashup
            foreach (var formula in formulas)
            {
                // Display the name of the query
                Console.WriteLine($"Query Name: {formula.Name}");

                // Attempt to obtain the service endpoint URL via reflection
                var urlProperty = formula.GetType().GetProperty("Url");
                if (urlProperty != null)
                {
                    string url = urlProperty.GetValue(formula) as string;
                    Console.WriteLine($"Service Endpoint URL: {url ?? "null"}");
                }
                else
                {
                    Console.WriteLine("Service Endpoint URL: not available via Url property.");
                }
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
