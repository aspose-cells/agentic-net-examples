using System;
using Aspose.Cells;
using Aspose.Cells.QueryTables;

class Program
{
    static void Main()
    {
        // Load an existing workbook that contains Power Query (OData) connections
        string sourcePath = "input.xlsx";
        Workbook workbook = new Workbook(sourcePath);

        // Access the DataMashup object which holds Power Query formulas
        DataMashup mashup = workbook.DataMashup;
        if (mashup != null)
        {
            var formulas = mashup.PowerQueryFormulas;
            Console.WriteLine($"Number of Power Query formulas: {formulas.Count}");

            foreach (var formula in formulas)
            {
                // Display the name of the Power Query formula
                Console.WriteLine($"Formula Name: {formula.Name}");

                // Attempt to retrieve a service endpoint URL.
                // Some Power Query formulas expose a Url property; use reflection to avoid
                // compile‑time dependency on a property that may not exist in all versions.
                var urlProperty = formula.GetType().GetProperty("Url");
                if (urlProperty != null)
                {
                    string url = urlProperty.GetValue(formula) as string;
                    Console.WriteLine($"Service Endpoint URL: {url}");
                }
                else
                {
                    Console.WriteLine("Service Endpoint URL not available for this formula.");
                }
            }
        }
        else
        {
            Console.WriteLine("No DataMashup information found in the workbook.");
        }

        // Save the workbook (optional, demonstrates lifecycle compliance)
        workbook.Save("output.xlsx");
    }
}