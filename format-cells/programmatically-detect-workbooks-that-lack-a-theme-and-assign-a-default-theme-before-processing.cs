// Title: Detect missing Excel workbook theme and log fallback .thmx file with Aspose.Cells for .NET
// AI Prompts: Generate C# code that checks Workbook.Theme and writes a warning if it is empty while a default .thmx file exists. | Create a method using Aspose.Cells that verifies a workbook has a theme and logs that applying a .thmx theme is not supported in the current version. | Show how to load an Excel file, detect the absence of a theme, and safely save the workbook after processing with Aspose.Cells.
// Common Searches: Aspose.Cells how to determine if an Excel file has a theme applied | C# check Workbook.Theme property for missing theme Aspose.Cells | log message when default .thmx theme file is present but cannot be applied Aspose.Cells | detect absent theme in .xlsx using Aspose.Cells .NET | fallback theme handling for Excel workbooks with Aspose.Cells
// Tags: Workbook.Theme property check Aspose.Cells | fallback .thmx theme handling .NET | detect missing Excel theme Aspose.Cells | log theme absence Aspose.Cells C# | theme support limitation Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook with Aspose.Cells, checks the Workbook.Theme property to see if a theme is assigned, and if none is found while a default .thmx file exists, it logs that applying a theme programmatically is not supported in the current API version. The workbook is then saved.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the workbook to process
            string workbookPath = "input.xlsx";

            // Path to the default theme file (must be a .thmx file)
            string defaultThemePath = "DefaultTheme.thmx";

            // Ensure the input workbook exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Input workbook not found: {workbookPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Determine whether the workbook already has a theme (Theme returns the theme name)
            bool hasTheme = !string.IsNullOrEmpty(workbook.Theme);

            // If no theme is present and a default theme file exists, note that applying a theme
            // is not directly supported via a single API in this version of Aspose.Cells.
            if (!hasTheme && File.Exists(defaultThemePath))
            {
                Console.WriteLine("Default theme file found, but applying a theme programmatically is not supported in this API version.");
                // Place any custom logic here if you need to manipulate styles manually.
            }

            // Save the workbook after processing
            workbook.Save("output.xlsx");
            Console.WriteLine("Workbook saved as output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
