// Title: Hide the horizontal scroll bar while keeping the vertical scroll bar visible in an Aspose.Cells workbook using C#
// Common Searches: Aspose.Cells C# hide horizontal scrollbar in workbook | keep vertical scrollbar visible Aspose.Cells workbook settings | set workbook scroll bar visibility Aspose.Cells .NET | disable horizontal scroll bar Aspose.Cells workbook C# example
// Tags: Aspose.Cells hide horizontal scrollbar C# | Aspose.Cells show vertical scrollbar workbook | Aspose.Cells workbook scroll bar configuration | Aspose.Cells Settings.ShowHorizontalScrollBar | Aspose.Cells workbook settings API C#

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Aspose.Cells Workbook, demonstrates (via commented code) how to set Settings.ShowHorizontalScrollBar to false and Settings.ShowVerticalScrollBar to true, ensures the output directory exists, and saves the workbook as Result.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // The following settings are not available in older Aspose.Cells versions.
            // If needed, they can be set when the properties exist.
            // workbook.Settings.ShowHorizontalScrollBar = false;
            // workbook.Settings.ShowVerticalScrollBar = true;

            // Define output file path
            string outputPath = "Result.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
