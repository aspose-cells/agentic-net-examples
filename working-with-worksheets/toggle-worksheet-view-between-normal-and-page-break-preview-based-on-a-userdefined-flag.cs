// Title: How to toggle an Excel worksheet between Normal view and Page Break Preview using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set worksheet to PageBreakPreview mode programmatically | switch Excel sheet view to normal view with Aspose.Cells .NET | toggle worksheet view based on a boolean flag using Aspose.Cells | change view type of a worksheet to PageBreakPreview in C# | Aspose.Cells ViewType enum example for Excel workbook
// Tags: Aspose.Cells ViewType enum | set worksheet view mode C# | page break preview Excel Aspose.Cells | toggle worksheet view programmatically | save workbook after view change Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing Excel file or creates a new workbook, accesses the first worksheet, and sets its ViewType to either PageBreakPreview or Normal based on a boolean flag, then saves the workbook to the specified output path.
    class Program
    {
        static void Main(string[] args)
        {
            // User‑defined flag: true = Page Break Preview, false = Normal view
            bool showPageBreakPreview = true;

            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a new one
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    // Create a new workbook with a single worksheet
                    workbook = new Workbook();
                    workbook.Worksheets[0].Name = "Sheet1";
                }

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Toggle the worksheet view based on the flag
                if (showPageBreakPreview)
                {
                    sheet.ViewType = ViewType.PageBreakPreview;
                }
                else
                {
                    // Normal view (value 0) – using explicit cast to avoid enum version issues
                    sheet.ViewType = (ViewType)0; // ViewType.Normal
                }

                // Ensure the output directory exists
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
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
