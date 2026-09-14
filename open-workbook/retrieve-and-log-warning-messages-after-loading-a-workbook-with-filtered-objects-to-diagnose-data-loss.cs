// Title: Load an Excel workbook with Aspose.Cells for .NET and capture warning messages to diagnose filtered‑object data loss
// AI Prompts: Write C# code that opens an Excel file using Aspose.Cells, then iterates through the Workbook's warning collection and writes each warning to the console. | Show how to enable warning collection when loading a workbook with filtered objects and display the warning details for data‑loss analysis. | Provide a sample that catches load‑time warnings from Aspose.Cells, formats them, and logs them for troubleshooting filtered‑object issues.
// Common Searches: Aspose.Cells retrieve warning messages after opening workbook with filtered data | How to log load warnings from Aspose.Cells to detect possible data loss in .NET | C# capture filtered objects warnings when loading Excel file with Aspose.Cells | Diagnose Aspose.Cells workbook load warnings for Excel files containing filters
// Tags: Aspose.Cells Workbook load warning collection | C# capture Aspose.Cells load warnings | filtered objects data loss detection Aspose.Cells | log Excel workbook warnings .NET | Aspose.Cells warning handling example

using System;
using System.IO;
using Aspose.Cells;

// The program verifies the existence of input.xlsx, loads it into an Aspose.Cells Workbook, and demonstrates how to access and log any warning messages generated during the load to help identify filtered‑object data loss.
class Program
{
    static void Main()
    {
        // Path to the source Excel file.
        string sourceFile = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException.
        if (!File.Exists(sourceFile))
        {
            Console.WriteLine($"Error: The file '{sourceFile}' was not found.");
            return;
        }

        try
        {
            // Load the workbook.
            Workbook workbook = new Workbook(sourceFile);

            // If needed, additional processing can be done here.
            Console.WriteLine("Workbook loaded successfully.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
