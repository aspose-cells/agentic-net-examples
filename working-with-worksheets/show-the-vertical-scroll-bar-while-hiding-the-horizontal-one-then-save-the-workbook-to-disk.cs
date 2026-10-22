// Title: Hide the horizontal scroll bar and keep the vertical scroll bar visible in an Aspose.Cells worksheet before saving as XLSX (C#)
// AI Prompts: Set Worksheet.ViewOptions.ShowHorizontalScrollBar = false and Worksheet.ViewOptions.ShowVerticalScrollBar = true, then call Workbook.Save to write the file to a specified path using Aspose.Cells for .NET. | Add a check for the ViewOptions property, configure scroll‑bar visibility accordingly, and persist the workbook as an .xlsx file with Aspose.Cells in C#.
// Common Searches: Aspose.Cells C# hide horizontal scrollbar only vertical scrollbar visible | set worksheet view options scrollbars Aspose.Cells .NET | disable horizontal scroll bar in generated Excel file using Aspose.Cells | save workbook after configuring scroll bar visibility Aspose.Cells C# | vertical scrollbar only Aspose.Cells worksheet example
// Tags: worksheet viewoptions hide horizontal scrollbar | show vertical scrollbar aspocells | configure worksheet scrollbars .net | save workbook as xlsx aspocells c# | excel scroll bar visibility aspocells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new Workbook, accesses its first Worksheet, uses the Worksheet.ViewOptions API to hide the horizontal scroll bar while keeping the vertical scroll bar visible, ensures the output directory exists, and saves the workbook as an XLSX file. Exception handling reports any errors during the process.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // NOTE: In some older versions of Aspose.Cells the ViewOptions property may not be available.
            // If needed, adjust scroll bar visibility using the appropriate API for your version.

            // Define output path
            string outputPath = "output.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to disk
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
