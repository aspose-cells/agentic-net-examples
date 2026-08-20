// Title: Refresh All Data Connections and Freeze Header Row Using Aspose.Cells for .NET (C#)
// Description: C# example that loads an Excel workbook with Aspose.Cells, updates every data connection (pivot tables, charts, external queries) via Workbook.RefreshAll, freezes the first row of the first worksheet using Worksheet.FreezePanes, and saves the refreshed file.
// Keywords: Aspose.Cells | Workbook.RefreshAll | Worksheet.FreezePanes | C# Excel automation | .NET Excel API | refresh data connections | freeze header row | pivot table refresh | external data source update | Excel view stability
// Common Searches: Aspose.Cells refresh all connections then freeze top row | C# Workbook.RefreshAll example | How to freeze header row after RefreshAll in Aspose.Cells | Freeze panes after updating data sources with Aspose.Cells | Refresh external links and lock first row in Excel using .NET
// Developer Intent: Refresh every data connection in a workbook and then lock the top row for stable viewing.
// Use Cases: Generate a final report where all linked data sources are up‑to‑date before distribution. | Keep column headers visible while users scroll through refreshed pivot tables or charts. | Automate Excel file preparation for dashboards that require both data refresh and consistent layout.
// AI Prompts: Create C# code with Aspose.Cells that calls RefreshAll and then freezes the first two rows of the first worksheet. | Explain the four‑parameter overload of Worksheet.FreezePanes and its best practice after a data refresh. | Show how to handle missing input files gracefully when refreshing connections and freezing panes with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // C# example that loads an Excel workbook with Aspose.Cells, updates every data connection (pivot tables, charts, external queries) via Workbook.RefreshAll, freezes the first row of the first worksheet using Worksheet.FreezePanes, and saves the refreshed file.
    public class RefreshAndFreezeDemo
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Refresh all data connections (pivot tables, charts, etc.)
                workbook.RefreshAll();

                // Freeze the top row in the first worksheet
                Worksheet firstSheet = workbook.Worksheets[0];
                // Freeze rows above row 1 (i.e., the first row). Use the 4‑parameter overload.
                firstSheet.FreezePanes(1, 0, 0, 0);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
