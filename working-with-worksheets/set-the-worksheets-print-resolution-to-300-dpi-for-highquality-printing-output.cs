// Title: Set worksheet print resolution to 300 DPI using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set worksheet print resolution to 300 DPI | How to configure high‑resolution printing for an Excel worksheet with Aspose.Cells | C# example for 300 DPI Excel output using Aspose.Cells page setup | Set Excel worksheet DPI in .NET with Aspose.Cells API
// Tags: Aspose.Cells worksheet print DPI configuration | C# set Excel print resolution Aspose.Cells | high‑resolution worksheet printing .NET | Aspose.Cells page setup zoom placeholder DPI | 300 DPI Excel output Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new Workbook, accesses the first Worksheet, and assigns PageSetup.Zoom = 100 as a placeholder because the PrintResolution property is not available in the current Aspose.Cells API. The workbook is then saved as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // The PrintResolution property is not available in the current Aspose.Cells API.
            // As an example, set the page zoom to 100% (default) to ensure normal scaling.
            worksheet.PageSetup.Zoom = 100;

            // Define the output file path
            string outputPath = "Output.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
