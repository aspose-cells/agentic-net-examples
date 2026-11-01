// Title: Mark worksheets that contain only shapes (no cell data) by setting the tab color to red using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that scans each worksheet, detects when Cells.MaxDataRow equals -1 and Shapes.Count is greater than zero, and changes the worksheet TabColor to red. | Create a .NET routine that loads an Excel workbook, flags sheets that have no cell data but contain drawings by applying a red tab color, and saves the updated file.
// Common Searches: how to highlight Excel worksheets that only have drawings using Aspose.Cells C# | Aspose.Cells C# detect worksheets with no cell data but with shapes | set worksheet tab color based on shape count Aspose.Cells .NET | flag shape‑only sheets in a workbook using Aspose.Cells API | identify empty data rows and shapes in Excel worksheets with Aspose.Cells
// Tags: detect shape‑only worksheets Aspose.Cells | set worksheet tab color C# Aspose.Cells | check MaxDataRow empty sheet Aspose.Cells | use Shapes.Count to find drawings Aspose.Cells | flag Excel sheets by tab color .NET

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

// // Loads an Excel file, iterates through each worksheet, and if a sheet has no cell data (Cells.MaxDataRow == -1) but contains one or more shapes (Shapes.Count > 0), it sets the worksheet's tab color to red before saving the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Examine each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // MaxDataRow == -1 indicates no cell data
                // Shapes.Count > 0 indicates shapes are present
                if (sheet.Cells.MaxDataRow == -1 && sheet.Shapes.Count > 0)
                {
                    // Flag the worksheet by setting its tab color to red
                    sheet.TabColor = Color.Red;
                }
            }

            // Save the workbook with flags applied
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
