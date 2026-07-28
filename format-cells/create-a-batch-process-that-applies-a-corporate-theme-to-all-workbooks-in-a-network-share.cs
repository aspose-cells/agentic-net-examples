// Title: C# – Batch apply a corporate Excel theme to all workbooks on a network share with Aspose.Cells
// Description: Loads a template workbook that contains the corporate theme, scans a UNC path for .xlsx, .xls and .xlsm files (including subfolders), copies the theme to each workbook using the CopyTheme method, overwrites the original file, and logs successes or errors.
// Keywords: Aspose.Cells | CopyTheme | C# | batch Excel theme | network share | UNC path | apply corporate theme | bulk Excel processing | automate workbook branding | Excel file enumeration
// Common Searches: apply Excel theme to multiple files Aspose.Cells | copy theme from template workbook C# | batch update Excel workbooks on network share | Aspose.Cells CopyTheme example | automate corporate branding in Excel files
// Developer Intent: Automatically apply a corporate Excel theme to every workbook stored on a shared network folder in one operation.
// Use Cases: Enforce brand‑consistent styling across all financial reports saved on a shared drive. | Refresh the visual design of archived spreadsheets after a company rebrand. | Run a nightly job that guarantees newly added workbooks adopt the latest corporate theme.
// AI Prompts: Write C# code that uses Aspose.Cells to copy a theme from a template workbook to all .xlsx, .xls, and .xlsm files in a UNC directory, with error handling and logging. | Refactor the batch theme‑application script to process files in parallel while keeping Aspose.Cells usage thread‑safe. | Explain how to add a backup step that saves each original workbook to a separate folder before applying the corporate theme.

using System;
using System.IO;
using Aspose.Cells;

// Loads a template workbook that contains the corporate theme, scans a UNC path for .xlsx, .xls and .xlsm files (including subfolders), copies the theme to each workbook using the CopyTheme method, overwrites the original file, and logs successes or errors.
class ApplyCorporateThemeBatch
{
    static void Main()
    {
        // Path to the corporate theme template workbook (contains the desired theme)
        string themeTemplatePath = @"\\Server\Share\CorporateThemeTemplate.xlsx";

        // Path to the network share that holds the workbooks to be processed
        string workbooksFolder = @"\\Server\Share\Workbooks";

        // Verify that the theme template exists
        if (!File.Exists(themeTemplatePath))
        {
            Console.WriteLine($"Theme template not found: {themeTemplatePath}");
            return;
        }

        // Verify that the workbooks folder exists
        if (!Directory.Exists(workbooksFolder))
        {
            Console.WriteLine($"Workbooks folder not found: {workbooksFolder}");
            return;
        }

        try
        {
            // Load the source workbook that carries the corporate theme
            using (Workbook themeWorkbook = new Workbook(themeTemplatePath))
            {
                // Retrieve all Excel files (xlsx, xls, xlsm) from the folder and its subfolders
                string[] files = Directory.GetFiles(workbooksFolder, "*.*", SearchOption.AllDirectories);
                foreach (string filePath in files)
                {
                    string ext = Path.GetExtension(filePath).ToLowerInvariant();
                    if (ext != ".xlsx" && ext != ".xls" && ext != ".xlsm")
                        continue; // Skip non‑Excel files

                    // Ensure the target file still exists before processing
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found (skipped): {filePath}");
                        continue;
                    }

                    try
                    {
                        // Load the target workbook
                        using (Workbook targetWorkbook = new Workbook(filePath))
                        {
                            // Apply the corporate theme from the template workbook
                            targetWorkbook.CopyTheme(themeWorkbook);

                            // Save the workbook, overwriting the original file
                            targetWorkbook.Save(filePath);
                        }

                        Console.WriteLine($"Theme applied: {filePath}");
                    }
                    catch (Exception exFile)
                    {
                        Console.WriteLine($"Error processing file '{filePath}': {exFile.Message}");
                    }
                }
            }

            Console.WriteLine("Corporate theme applied to all workbooks successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
