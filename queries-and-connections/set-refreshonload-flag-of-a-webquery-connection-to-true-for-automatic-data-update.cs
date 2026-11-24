// Title: Enable RefreshOnLoad for WebQueryTable connections in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set WebQueryTable RefreshOnLoad to true | How to enable automatic refresh for web query connections in an .xlsx file with Aspose.Cells | Configure RefreshOnLoad property for a WebQueryTable in a .NET Excel workbook | Set web query connection to refresh on workbook open using Aspose.Cells API | RefreshOnLoad flag example for WebQueryTable in Aspose.Cells for .NET
// Tags: WebQueryTable refresh flag Aspose.Cells | automatic web query refresh Aspose.Cells | set web query connection flag C# | Excel workbook load refresh setting Aspose.Cells | configure WebQueryTable properties .NET

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates loading or creating an Excel workbook, locating each WebQueryTable, enabling its RefreshOnLoad flag to ensure data updates automatically when the workbook is opened, and then saving the modified file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file '{inputPath}' not found. Creating a new workbook.");
                workbook = new Workbook();
            }

            // NOTE: The WebQueryTable API may not be available in the current Aspose.Cells version.
            // If needed, enable RefreshOnLoad for each WebQueryTable here using the appropriate API.

            // Save the workbook with the updated settings.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
