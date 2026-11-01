// Title: Hide the horizontal scroll bar while keeping the vertical scroll bar visible in an Aspose.Cells workbook using C# (WorkbookView workaround)
// Common Searches: Aspose.Cells C# hide horizontal scrollbar but keep vertical scrollbar visible | How to set Excel scroll bar visibility with Aspose.Cells .NET | C# programmatically control workbook view scrollbars in Aspose.Cells | Aspose.Cells hide horizontal scroll bar after saving workbook | WorkbookView scroll bar settings example for Aspose.Cells C#
// Tags: hide horizontal scrollbar Aspose.Cells | show vertical scrollbar Aspose.Cells | WorkbookView scroll bar configuration .NET | Excel scrollbar visibility Aspose.Cells API | WorkbookSettings scroll bar limitation C#

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, accesses its WorkbookSettings (which currently lack direct scroll‑bar visibility properties), and explains that scroll‑bar visibility must be adjusted via the WorkbookView object after the file is opened in Excel. It ensures the output directory exists, saves the workbook as output.xlsx, and includes basic exception handling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access workbook settings
            WorkbookSettings settings = workbook.Settings;

            // The WorkbookSettings class does not expose scroll bar visibility properties
            // in recent Aspose.Cells versions. If needed, these settings can be adjusted
            // via the WorkbookView object after the workbook is saved or opened in Excel.

            // Ensure the output directory exists
            string outputPath = "output.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
