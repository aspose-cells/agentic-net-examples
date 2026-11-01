// Title: How to toggle an Aspose.Cells worksheet between Normal view and Page Break Preview using a boolean flag in C#
// AI Prompts: Generate C# code that reads a boolean variable and sets the first worksheet's ViewType to PageBreakPreview when true or to Normal when false, then saves the workbook. | Create a reusable C# method that accepts a Worksheet object and a bool flag to switch its view mode between Normal and Page Break Preview using Aspose.Cells.
// Common Searches: Aspose.Cells C# change worksheet view to page break preview based on condition | programmatically set Excel sheet view mode to normal or page break preview with Aspose.Cells | toggle worksheet ViewType enum in .NET using a flag | how to save workbook after changing view type in Aspose.Cells | C# example for switching between normal view and page break preview in Excel file
// Tags: Aspose.Cells set worksheet view type | C# toggle page break preview mode | ViewType enum normal view Aspose.Cells | save workbook after view change Aspose.Cells | conditional worksheet view switching .NET

using Aspose.Cells;
using System;
using System.IO;

// The example creates or loads a workbook, checks a boolean flag, assigns the worksheet's ViewType to PageBreakPreview or Normal accordingly, ensures the output directory exists, saves the workbook as output.xlsx, and handles any errors.
class Program
{
    static void Main(string[] args)
    {
        try
        {
            // Flag to choose view mode: true = Page Break Preview, false = Normal view
            bool usePageBreakPreview = true; // Replace with actual flag source as needed

            // Create a new workbook (or load an existing one if needed)
            Workbook workbook = new Workbook(); // For loading: new Workbook("input.xlsx");

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set the worksheet view based on the flag
            if (usePageBreakPreview)
            {
                sheet.ViewType = ViewType.PageBreakPreview;
            }
            else
            {
                // Normal view is the default; cast to the underlying enum value if needed
                sheet.ViewType = (ViewType)0; // Normal view
            }

            // Define output file path
            string outputPath = "output.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? Directory.GetCurrentDirectory();
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the updated view setting
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
