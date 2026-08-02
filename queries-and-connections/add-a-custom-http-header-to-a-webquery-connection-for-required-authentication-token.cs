using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class AddCustomHeaderDemo
{
    static void Main()
    {
        try
        {
            const string inputPath = "InputWithWebQuery.xlsx";
            const string outputPath = "OutputWithCustomHeader.xlsx";

            // Ensure the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook that contains a WebQuery connection
            Workbook workbook = new Workbook(inputPath);

            // Locate the first WebQueryConnection in the workbook
            WebQueryConnection webConn = null;
            foreach (ExternalConnection conn in workbook.DataConnections)
            {
                if (conn is WebQueryConnection wqc)
                {
                    webConn = wqc;
                    break;
                }
            }

            if (webConn == null)
            {
                Console.WriteLine("No WebQueryConnection found in the workbook.");
                return;
            }

            // Set the URL of the web query (the endpoint that returns the data)
            webConn.Url = "https://api.example.com/data";

            // Add a custom HTTP header for authentication (using Post as a workaround)
            string authToken = "your_auth_token_here";
            webConn.Post = $"Authorization: Bearer {authToken}";

            // Optional: configure other properties as needed
            webConn.IsHtmlTables = true;

            // Save the workbook with the modified connection
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}