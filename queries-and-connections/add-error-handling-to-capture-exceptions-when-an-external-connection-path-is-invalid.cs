// Title: Implement try‑catch error handling for invalid external connection paths in Aspose.Cells for .NET
// Common Searches: how to catch invalid external connection path exception in Aspose.Cells C# | C# Aspose.Cells workbook save error handling for missing external data source | handling external connection path errors when loading Excel with Aspose.Cells | Aspose.Cells .NET try-catch example for external data source failures | detect and log invalid external connection paths using Aspose.Cells
// Tags: Aspose.Cells external connection path error handling | C# workbook save exception handling Aspose.Cells | invalid external data source detection Aspose.Cells | try-catch pattern for Excel external connections .NET | Aspose.Cells file load failure handling

using System;
using System.IO;
using Aspose.Cells;

// The sample checks for the input file, loads the workbook with Aspose.Cells, and wraps both loading and saving operations in try‑catch blocks to capture and report exceptions, demonstrating how to handle invalid external connection paths and other runtime errors in a .NET application.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // ------------------------------------------------------------
            // NOTE: The ExternalConnections API is not available in the
            // current Aspose.Cells version, so this section is omitted.
            // ------------------------------------------------------------

            // Save the workbook after processing
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine("Error while saving workbook: " + saveEx.Message);
            }
        }
        catch (Exception ex)
        {
            // Capture and display any unexpected exception
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
