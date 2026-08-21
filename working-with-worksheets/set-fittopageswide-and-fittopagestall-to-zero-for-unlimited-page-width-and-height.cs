// Title: Aspose.Cells for .NET – Set FitToPagesWide & FitToPagesTall to 0 for Unlimited Page Size
// Description: Shows how to configure a worksheet’s PageSetup in Aspose.Cells (.NET) by assigning 0 to FitToPagesWide and FitToPagesTall, disabling percent‑scale printing, and saving the workbook, allowing unlimited page width and height.
// Keywords: Aspose.Cells | .NET | C# | FitToPagesWide | FitToPagesTall | unlimited pages | page setup | disable percent scaling | Excel printing | worksheet layout | Aspose.Cells API | zero page scaling
// Common Searches: Aspose.Cells set FitToPagesWide to 0 | FitToPagesTall zero Aspose.Cells | unlimited page size Aspose.Cells .NET | disable percent scale Aspose.Cells | print Excel without page limits using Aspose | Aspose.Cells page setup unlimited | C# Aspose.Cells page scaling | remove page scaling Aspose.Cells
// Developer Intent: Configure worksheet printing to use unlimited pages by setting FitToPagesWide and FitToPagesTall to zero and turning off percent scaling.
// Use Cases: Create a new workbook and set the first worksheet to print without page‑size constraints before saving. | Adjust an existing Excel file’s page setup so content can span any number of pages for custom report layouts. | Generate Excel output for large tables or charts where automatic scaling would truncate or shrink data.
// AI Prompts: Provide C# code using Aspose.Cells to set FitToPagesWide = 0 and FitToPagesTall = 0 and disable percent scaling for a worksheet. | Explain the effect of setting FitToPagesWide and FitToPagesTall to zero on Excel printing behavior in Aspose.Cells. | Give a step‑by‑step tutorial for creating a workbook, configuring unlimited page dimensions, and saving it with Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to configure a worksheet’s PageSetup in Aspose.Cells (.NET) by assigning 0 to FitToPagesWide and FitToPagesTall, disabling percent‑scale printing, and saving the workbook, allowing unlimited page width and height.
    public class FitToPagesUnlimitedDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Set FitToPagesWide and FitToPagesTall to 0 (unlimited width and height)
                worksheet.PageSetup.FitToPagesWide = 0;
                worksheet.PageSetup.FitToPagesTall = 0;

                // Ensure the FitToPages settings are used instead of percent scaling
                worksheet.PageSetup.IsPercentScale = false;

                // Define output file path
                string outputPath = "FitToPagesUnlimited.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            FitToPagesUnlimitedDemo.Run();
        }
    }
}
