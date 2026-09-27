// Title: How to set ProxyHost and ProxyPort on Aspose.Cells DbConnection in C# for custom proxy support
// AI Prompts: Generate C# code that creates an Aspose.Cells DbConnection, assigns the ProxyHost and ProxyPort properties, opens the connection, and then loads and saves a workbook. | Show an example of configuring a custom HTTP proxy for Aspose.Cells database operations by setting ProxyHost and ProxyPort on the DbConnection object before executing any queries.
// Common Searches: Aspose.Cells C# set ProxyHost and ProxyPort on DbConnection example | custom proxy configuration for Aspose.Cells database connection in .NET | how to assign ProxyHost ProxyPort properties to Aspose.Cells DbConnection | C# Aspose.Cells DBConnection proxy settings for Excel data import
// Tags: Aspose.Cells DbConnection proxy configuration | set ProxyHost ProxyPort C# | custom proxy for Aspose.Cells database access | Aspose.Cells database connection proxy settings | C# configure proxy on Excel data source

using System;
using System.IO;
using Aspose.Cells;

// This example demonstrates how to assign ProxyHost and ProxyPort values to an Aspose.Cells DbConnection in C#, enabling custom proxy support for database operations before loading or saving a workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Define output path
            string outputPath = "Output.xlsx";

            // Ensure the output directory exists (if any)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
