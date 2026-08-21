// Title: Extract OData Endpoint URLs from DataMashup Power Query Formulas with Aspose.Cells for .NET
// Description: Shows how to load an Excel workbook using Aspose.Cells, access its DataMashup object, iterate through Power Query formulas, and retrieve the OData service endpoint via reflection (ConnectionString or Url). The endpoints are written to the console, with optional workbook saving.
// Keywords: Aspose.Cells | C# | .NET | DataMashup | Power Query | OData endpoint | ConnectionString | Url property | reflection | extract OData URL | Excel workbook | query tables | API example
// Common Searches: Aspose.Cells get OData URL from DataMashup | Read Power Query connection string Aspose.Cells | C# extract OData service endpoint Excel | How to list OData connections in workbook using Aspose | Retrieve DataMashup formulas Aspose.Cells | Get OData endpoint from Power Query formula .NET
// Developer Intent: Retrieve and log the OData service endpoint URLs defined in a workbook’s DataMashup Power Query formulas.
// Use Cases: Validate external OData source URLs before refreshing Power Query connections. | Audit OData connections across multiple workbooks for compliance reporting. | Generate a summary of Power Query formulas and their endpoints for documentation. | Automate health checks of external data sources in batch processing pipelines.
// AI Prompts: Write a reusable method that returns a list of OData endpoint strings from a Workbook’s DataMashup using Aspose.Cells. | Enhance the sample to also capture endpoints stored in a custom property named 'ServiceUrl' on the formula objects. | Create an example that writes the extracted OData endpoints to a CSV file instead of the console. | Generate unit tests that verify the reflection logic correctly extracts ConnectionString and Url values.

using System;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.QueryTables;

// Shows how to load an Excel workbook using Aspose.Cells, access its DataMashup object, iterate through Power Query formulas, and retrieve the OData service endpoint via reflection (ConnectionString or Url). The endpoints are written to the console, with optional workbook saving.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the DataMashup object which holds Power Query formulas
        DataMashup mashup = workbook.DataMashup;

        if (mashup != null)
        {
            // Iterate through all Power Query formulas
            foreach (var formula in mashup.PowerQueryFormulas)
            {
                // Attempt to retrieve the OData service endpoint.
                // Some versions expose it via ConnectionString, others via Url.
                string endpoint = string.Empty;

                Type formulaType = formula.GetType();

                // Try ConnectionString property
                PropertyInfo connProp = formulaType.GetProperty("ConnectionString");
                if (connProp != null)
                {
                    endpoint = connProp.GetValue(formula) as string;
                }
                else
                {
                    // Fallback to Url property
                    PropertyInfo urlProp = formulaType.GetProperty("Url");
                    if (urlProp != null)
                    {
                        endpoint = urlProp.GetValue(formula) as string;
                    }
                }

                // Log the details
                Console.WriteLine($"Power Query Formula: {formula.Name}");
                Console.WriteLine($"OData Service Endpoint: {endpoint}");
                Console.WriteLine(new string('-', 40));
            }
        }
        else
        {
            Console.WriteLine("No DataMashup information found in the workbook.");
        }

        // Save the workbook (optional, adjust path as needed)
        workbook.Save("output.xlsx");
    }
}
