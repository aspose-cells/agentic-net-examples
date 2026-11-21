// Title: Enable Page Break Preview view and print gridlines on an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set worksheet view to PageBreakPreview | How to display printed page divisions in Excel with Aspose.Cells | Enable print gridlines while using page break preview in Aspose.Cells | Saving an Excel file with page break preview mode using Aspose.Cells for .NET | Aspose.Cells example to configure worksheet viewtype in C#
// Tags: Aspose.Cells worksheet viewtype PageBreakPreview | Aspose.Cells print gridlines configuration | Aspose.Cells save workbook with page layout preview | Aspose.Cells C# worksheet view settings | Aspose.Cells Excel page break visualization

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new Workbook (or loads an existing one), accesses the first Worksheet, sets its ViewType to PageBreakPreview to visualize printed page divisions, enables PrintGridlines in the PageSetup, ensures the output directory exists, and saves the workbook as "PageBreakPreview.xlsx".
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing file with new Workbook("input.xlsx"))
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Enable Page Break Preview mode to visualize printed page divisions
            sheet.ViewType = ViewType.PageBreakPreview;

            // Show gridlines in the printed output
            sheet.PageSetup.PrintGridlines = true;

            // Define output file path
            string outputPath = "PageBreakPreview.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
