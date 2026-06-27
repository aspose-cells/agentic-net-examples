using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string templatePath = "TemplateWithOData.xlsx";
                const string resultPath = "Result.xlsx";

                // Verify that the template file exists.
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file not found: {templatePath}");
                    return;
                }

                // Load the workbook that already contains an OData connection.
                Workbook workbook = new Workbook(templatePath);

                // Ensure the workbook has at least one external data connection.
                if (workbook.DataConnections.Count == 0)
                {
                    Console.WriteLine("No external data connections found in the workbook.");
                    return;
                }

                // Retrieve the first external connection (assumed to be the OData connection).
                ExternalConnection odataConnection = workbook.DataConnections[0];

                // Get the first worksheet where the query table will be placed.
                Worksheet worksheet = workbook.Worksheets[0];

                // OPTIONAL: Refresh the OData connection to pull latest data.
                // odataConnection.Refresh();

                // Save the workbook to the result file.
                workbook.Save(resultPath);
                Console.WriteLine($"Workbook saved successfully to {resultPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}