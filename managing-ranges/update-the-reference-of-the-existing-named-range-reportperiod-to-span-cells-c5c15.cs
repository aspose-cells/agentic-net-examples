// Title: Aspose.Cells for .NET – Update the "ReportPeriod" named range to cells C5:C15
// Description: Loads an existing workbook, finds the named range "ReportPeriod", sets its RefersTo property to =$SheetName!$C$5:$C$15, and saves the updated file. Includes checks for missing files and absent ranges.
// Keywords: Aspose.Cells | C# | .NET | named range | update RefersTo | ReportPeriod | C5:C15 | Excel automation | modify named range | Aspose.Cells example
// Common Searches: Aspose.Cells change named range reference C# | Update existing named range in Excel using Aspose.Cells | .NET set RefersTo property for a named range | How to modify named range ReportPeriod Aspose.Cells | Set named range to C5:C15 with Aspose.Cells
// Developer Intent: Change the RefersTo address of the existing ReportPeriod named range so it covers cells C5 through C15 on its worksheet.
// Use Cases: Adjust a reporting period range after inserting or deleting rows in the source data. | Align a template's named range with a newly calculated data block before exporting to downstream systems. | Standardize a fixed column segment for batch processing or chart data sources. | Correct a mis‑defined range in a legacy workbook without manual Excel editing.
// AI Prompts: Show code that verifies the ReportPeriod named range now points to C5:C15 after saving. | Generate a fallback routine that creates ReportPeriod if it does not exist, then sets its reference to C5:C15. | Explain how to update a named range located on a different worksheet using Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

namespace UpdateNamedRangeApp
{
    // Loads an existing workbook, finds the named range "ReportPeriod", sets its RefersTo property to =$SheetName!$C$5:$C$15, and saves the updated file. Includes checks for missing files and absent ranges.
    class UpdateNamedRange
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (assumes the named range is on this sheet)
                Worksheet sheet = workbook.Worksheets[0];

                // Retrieve the existing named range "ReportPeriod"
                Name reportPeriod = workbook.Worksheets.Names["ReportPeriod"];
                if (reportPeriod == null)
                {
                    Console.WriteLine("Named range 'ReportPeriod' not found.");
                    return;
                }

                // Update the reference to span cells C5:C15 on the same worksheet
                reportPeriod.RefersTo = $"={sheet.Name}!$C$5:$C$15";

                // Save the workbook with the updated named range
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
