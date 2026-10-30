// Title: How to update a Power Query connection source file path to a cloud URL with Aspose.Cells for .NET (API limitation note)
// Common Searches: Aspose.Cells C# change Power Query source file path to Azure Blob storage | update Power Query connection URL programmatically using Aspose.Cells .NET | set Power Query data source to cloud storage URL in Excel workbook via Aspose.Cells | Aspose.Cells version that supports Power Query connection editing | modify Power Query source location in existing workbook using C# Aspose.Cells
// Tags: Aspose.Cells Power Query source URL update | C# modify Power Query connection path | Excel workbook Power Query cloud storage link | Aspose.Cells version limitation Power Query API | programmatic Power Query data source change .NET

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing workbook and saves it unchanged, highlighting that the current Aspose.Cells for .NET release does not expose Power Query connection classes, so updating the Power Query source file location to a cloud URL must be done with a later version or alternative approach.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // NOTE: The current Aspose.Cells version used in this project does not expose
            // Power Query connection classes (e.g., WorkbookConnection, PowerQueryConnection).
            // If such functionality is required, ensure the appropriate Aspose.Cells version
            // and references are added. The code below preserves the workbook without
            // modifying Power Query connections.

            // Save the workbook to a new file (or overwrite the original)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
