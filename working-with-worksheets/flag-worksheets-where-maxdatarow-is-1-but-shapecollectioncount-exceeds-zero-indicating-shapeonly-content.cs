// Title: Identify and report worksheets that contain only shapes (no data rows) using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells find worksheets with no cell data but with drawings | How to list Excel sheets that have only shapes using Aspose.Cells | Detect empty data rows and shape collection in a workbook with Aspose.Cells .NET | Generate a report of shape‑only worksheets in an Excel file using Aspose.Cells C# | Flag worksheets where Cells.MaxDataRow is -1 and Shapes.Count > 0 in Aspose.Cells
// Tags: Aspose.Cells shape‑only worksheet detection | C# MaxDataRow -1 with Shapes.Count check | worksheet reporting for drawings without cell data | enumerate Excel sheets containing only drawings Aspose.Cells | generate shape‑only worksheet summary .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeOnlyWorksheetFlagger
{
    // The example loads an Excel workbook with Aspose.Cells, iterates each worksheet, and flags any sheet where Cells.MaxDataRow equals -1 (indicating no data rows) while the Shapes collection contains one or more items. It then writes a concise report to the console or an optional output file.
    class Program
    {
        /// <param name="args">Command‑line arguments.</param>
        static void Main(string[] args)
        {
            // Validate arguments
            if (args == null || args.Length == 0)
            {
                Console.WriteLine("Usage: ShapeOnlyWorksheetFlagger <inputWorkbookPath> [outputReportPath]");
                return;
            }

            string inputPath = args[0];
            string outputPath = args.Length > 1 ? args[1] : null;

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Prepare a StringWriter to collect the report (will be written to console or file later)
                StringWriter reportWriter = new StringWriter();

                reportWriter.WriteLine($"Workbook: {Path.GetFileName(inputPath)}");
                reportWriter.WriteLine($"Processed on: {DateTime.Now}");
                reportWriter.WriteLine(new string('=', 50));

                bool anyFlagged = false;

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // MaxDataRow returns -1 when there is no cell containing data
                    int maxDataRow = sheet.Cells.MaxDataRow;

                    // ShapeCollection.Count gives the number of shapes on the worksheet
                    int shapeCount = sheet.Shapes.Count;

                    // Flag condition: no data rows but at least one shape
                    if (maxDataRow == -1 && shapeCount > 0)
                    {
                        anyFlagged = true;
                        reportWriter.WriteLine($"Worksheet \"{sheet.Name}\" is shape‑only:");
                        reportWriter.WriteLine($"  MaxDataRow = {maxDataRow}");
                        reportWriter.WriteLine($"  Shape count = {shapeCount}");
                        reportWriter.WriteLine();
                    }
                }

                if (!anyFlagged)
                {
                    reportWriter.WriteLine("No worksheets with shape‑only content were found.");
                }

                // Output the report
                if (string.IsNullOrWhiteSpace(outputPath))
                {
                    // Write to console
                    Console.WriteLine(reportWriter.ToString());
                }
                else
                {
                    // Ensure the directory for the output file exists
                    string outputDir = Path.GetDirectoryName(outputPath);
                    if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    File.WriteAllText(outputPath, reportWriter.ToString());
                    Console.WriteLine($"Report written to \"{outputPath}\"");
                }
            }
            catch (Exception ex)
            {
                // Do not hide failures – report the full exception
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.ToString());
            }
        }
    }
}
