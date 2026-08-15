// Title: Batch apply a corporate Excel theme to all workbooks on a network share with Aspose.Cells for .NET
// Description: C# utility that loads a template workbook containing the corporate theme, enumerates every .xlsx file in a network share, copies the theme to each workbook via Aspose.Cells CopyTheme, and saves the files, with built‑in error handling and logging.
// Keywords: Aspose.Cells | CopyTheme | C# | .NET | Excel theme automation | batch Excel processing | network share | corporate branding | bulk workbook update | Excel styling script
// Common Searches: How to copy an Excel theme to multiple files using Aspose.Cells C# | Batch apply corporate theme to workbooks on a shared folder | Aspose.Cells CopyTheme example for network drives | C# script to update Excel themes in a directory | Automate Excel theme changes across many files
// Developer Intent: Programmatically apply a corporate Excel theme to every workbook stored on a shared network location.
// Use Cases: Enforce brand consistency across departmental reports saved on a shared drive. | Refresh the visual style of all existing Excel outputs after a rebranding initiative without manual editing. | Integrate into a CI/CD pipeline to guarantee that generated Excel files always use the corporate theme.
// AI Prompts: Generate C# code that uses Aspose.Cells to copy a theme from a template workbook to all .xlsx files in a specified folder, including robust error handling and progress logging. | Show how to add a progress bar or console output that reports the number of workbooks processed and any failures during a bulk theme update on a network share. | Provide a modification to skip files that already contain the target corporate theme, using Aspose.Cells metadata inspection.

using System;
using System.IO;
using Aspose.Cells;

namespace CorporateThemeBatch
{
    // C# utility that loads a template workbook containing the corporate theme, enumerates every .xlsx file in a network share, copies the theme to each workbook via Aspose.Cells CopyTheme, and saves the files, with built‑in error handling and logging.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook that contains the corporate theme (template file)
            string themeTemplatePath = @"\\NetworkShare\Templates\CorporateThemeTemplate.xlsx";

            // Verify the template exists
            if (!File.Exists(themeTemplatePath))
            {
                Console.WriteLine($"Theme template not found: {themeTemplatePath}");
                return;
            }

            // Load the source workbook that holds the desired theme
            Workbook sourceWorkbook = new Workbook(themeTemplatePath);

            // Path to the network share folder containing workbooks to process
            string workbooksFolder = @"\\NetworkShare\Workbooks";

            // Get all Excel files (you can adjust the search pattern as needed)
            string[] excelFiles = Directory.GetFiles(workbooksFolder, "*.xlsx", SearchOption.AllDirectories);

            foreach (string filePath in excelFiles)
            {
                try
                {
                    // Load the target workbook
                    Workbook targetWorkbook = new Workbook(filePath);

                    // Copy the corporate theme from the source workbook
                    targetWorkbook.CopyTheme(sourceWorkbook);

                    // Save the workbook, overwriting the original file
                    targetWorkbook.Save(filePath, SaveFormat.Xlsx);

                    Console.WriteLine($"Applied theme to: {filePath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch theme application completed.");
        }
    }
}
