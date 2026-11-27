// Title: Refresh all data connections and freeze the top row of each worksheet with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# refresh all workbook connections and freeze first row | C# code to refresh data connections then apply FreezePanes on each sheet using Aspose.Cells | How to recalculate formulas after refreshing connections in an Aspose.Cells workbook | Save workbook after refreshing connections and freezing panes with Aspose.Cells for .NET | Command‑line tool to refresh Excel data connections and freeze top rows using Aspose.Cells
// Tags: Workbook.RefreshAll Aspose.Cells | Worksheet.FreezePanes top row | recalculate formulas Aspose.Cells | save refreshed workbook C# | command line Aspose.Cells refresh connections

using System;
using System.IO;
using Aspose.Cells;

namespace RefreshAndFreezeDemo
{
    // The example loads a workbook, refreshes every data connection with Workbook.RefreshAll(), recalculates all formulas, freezes the first row of each worksheet using Worksheet.FreezePanes(1,0,1,0), and saves the updated file to a new location.
    class Program
    {
        static void Main(string[] args)
        {
            // Validate command‑line arguments
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: RefreshAndFreezeDemo <inputWorkbookPath> [outputWorkbookPath]");
                return;
            }

            string inputPath = args[0];
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
                return;
            }

            // Determine output path
            string outputPath;
            if (args.Length >= 2)
            {
                outputPath = args[1];
            }
            else
            {
                string dir = Path.GetDirectoryName(inputPath) ?? string.Empty;
                string nameWithoutExt = Path.GetFileNameWithoutExtension(inputPath);
                string ext = Path.GetExtension(inputPath);
                outputPath = Path.Combine(dir, $"{nameWithoutExt}_refreshed{ext}");
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Console.WriteLine($"Loaded workbook \"{inputPath}\" successfully.");

                // Refresh all data connections in the workbook
                Console.WriteLine("Refreshing all data connections...");
                workbook.RefreshAll(); // Correct API to refresh connections
                Console.WriteLine("Data connections refreshed.");

                // Recalculate formulas after refresh
                Console.WriteLine("Recalculating formulas...");
                workbook.CalculateFormula();
                Console.WriteLine("Formula recalculation completed.");

                // Freeze the top row in each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Freeze the first row (row index 0)
                    sheet.FreezePanes(1, 0, 1, 0);
                    Console.WriteLine($"Frozen top row in worksheet \"{sheet.Name}\".");
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during processing:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
