// Title: Apply FreezePanes to the first row only on worksheets whose names start with "Report" using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that iterates through all worksheets and freezes the top row only when the sheet name begins with "Report". | Modify an existing Aspose.Cells workbook routine to add conditional FreezePanes logic based on a worksheet name prefix and then save the file. | Create a C# method that creates the output directory if missing before saving a workbook after applying conditional freeze panes with Aspose.Cells.
// Common Searches: Aspose.Cells C# freeze first row only on sheets that start with a specific prefix | How to conditionally apply FreezePanes in a .NET workbook based on worksheet name | C# example for freezing panes on Excel worksheets named "Report*" using Aspose.Cells | Save Aspose.Cells workbook to a folder that may not exist, creating the directory automatically
// Tags: conditional FreezePanes based on worksheet name Aspose.Cells | freeze top row Excel sheet Aspose.Cells C# | iterate worksheets and apply FreezePanes Aspose.Cells | ensure output directory exists before saving workbook C# | load or create workbook with Aspose.Cells .NET

using Aspose.Cells;
using System;
using System.IO;

// The example loads (or creates) a workbook, iterates through its worksheets, freezes the first row on any sheet whose name begins with "Report", ensures the output folder exists, and saves the workbook to output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook or create a new one if the file is missing
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook with one worksheet
                workbook.Worksheets[0].Name = "Report1"; // sample sheet name to trigger freezing
            }

            // Apply freeze panes to worksheets whose name starts with "Report"
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.Name.StartsWith("Report", StringComparison.OrdinalIgnoreCase))
                {
                    // Freeze the first row (row index 1, column index 0), freezing 1 row and 0 columns
                    sheet.FreezePanes(1, 0, 1, 0);
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
