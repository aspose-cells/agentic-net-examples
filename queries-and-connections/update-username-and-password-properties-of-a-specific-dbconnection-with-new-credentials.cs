// Title: How to programmatically change the UserName and Password of a specific DBConnection in an Excel workbook using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells update DataConnection credentials for a named connection | set new username and password for Excel DB connection with Aspose.Cells | modify specific DBConnection properties in a workbook using Aspose.Cells .NET API | change Excel workbook data source login details via Aspose.Cells code | Aspose.Cells iterate DataConnections collection to edit credentials
// Tags: Aspose.Cells programmatic DBConnection credential update | C# edit Excel DataConnection UserName property | Aspose.Cells dynamic DataConnections enumeration | save workbook after modifying connection settings | handle missing input file with Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel file, searches its DataConnections collection for a connection named 'MyDbConnection', updates the UserName and Password properties with new values, ensures the output directory exists, and saves the workbook to a new file while handling missing input files and runtime exceptions.
class Program
{
    static void Main()
    {
        // Input and output file paths
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Target DB connection name and new credentials
        string targetConnectionName = "MyDbConnection";
        string newUserName = "newUser";
        string newPassword = "newPassword";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Locate the specific data connection using dynamic to avoid compile‑time type issues
            dynamic targetConnection = null;
            foreach (dynamic conn in workbook.DataConnections)
            {
                if (conn.Name == targetConnectionName)
                {
                    targetConnection = conn;
                    break;
                }
            }

            if (targetConnection != null)
            {
                // Update the credentials
                targetConnection.UserName = newUserName;
                targetConnection.Password = newPassword;
                Console.WriteLine($"Connection '{targetConnectionName}' credentials updated.");
            }
            else
            {
                Console.WriteLine($"Connection '{targetConnectionName}' not found.");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with updated connection information
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
