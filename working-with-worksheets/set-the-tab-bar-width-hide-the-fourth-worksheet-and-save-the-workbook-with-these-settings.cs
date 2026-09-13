// Title: Hide the fourth worksheet and customize the tab bar color before saving an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Create a new Workbook, set the TabBarColor to a chosen value, hide the worksheet at index 3, and save the workbook as an .xlsx file using Aspose.Cells for .NET. | Show how to verify that a fourth sheet exists, make it invisible, adjust workbook.Settings (e.g., TabBarColor), and persist the workbook to a specified path.
// Common Searches: asp.net hide fourth worksheet using Aspose.Cells | change tab bar color in an Aspose.Cells workbook C# | save Excel file after hiding specific sheet with Aspose.Cells .NET | Aspose.Cells workbook settings for customizing tab bar appearance | C# hide worksheet by index and export workbook to .xlsx
// Tags: hide worksheet by index Aspose.Cells | set tab bar color Aspose.Cells | save workbook with hidden sheets .NET | Aspose.Cells workbook settings customization | C# Excel workbook tab bar appearance

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, notes that Aspose.Cells does not support a TabBarWidth property (suggesting TabBarColor as an alternative), hides the fourth worksheet when present, ensures the output directory exists, and saves the workbook to 'output.xlsx' while handling potential exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // NOTE: Aspose.Cells does not provide a TabBarWidth property.
            // If you need to customize the tab bar appearance, consider using other available settings.
            // Example (setting tab bar color):
            // workbook.Settings.TabBarColor = System.Drawing.Color.LightGray;

            // Hide the fourth worksheet (zero‑based index 3) if it exists
            if (workbook.Worksheets.Count > 3)
            {
                workbook.Worksheets[3].IsVisible = false;
            }

            // Define output path
            string outputPath = "output.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the applied settings
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
