// Title: C# – Export each worksheet to its own HTML file using Aspose.Cells and reuse a single TableCssId
// Common Searches: how to batch export all worksheets to individual HTML files using Aspose.Cells in C# | Aspose.Cells C# export active worksheet only to HTML with shared TableCssId | create safe file names from worksheet names when saving as HTML in .NET | set TableCssId for multiple HTML outputs with Aspose.Cells HtmlSaveOptions | C# example for exporting each Excel sheet to separate HTML files using Aspose.Cells
// Tags: Aspose.Cells export worksheets to separate HTML files C# | shared TableCssId for multiple HTML exports Aspose.Cells | HtmlSaveOptions ExportActiveWorksheetOnly usage | safe filename generation for worksheet names .NET | batch HTML conversion of Excel workbook Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The C# program loads an Excel workbook, creates an output folder, and iterates through each worksheet. For every sheet it activates the worksheet, configures HtmlSaveOptions with ExportActiveWorksheetOnly=true and a common TableCssId, generates a filesystem‑safe file name from the sheet name, and saves the sheet as an individual HTML file, reusing the same CSS table identifier across all files.
class Program
{
    static void Main()
    {
        // Path to the source Excel file
        string excelPath = "input.xlsx";

        // Verify that the input file exists
        if (!File.Exists(excelPath))
        {
            Console.WriteLine($"Error: The file \"{excelPath}\" was not found.");
            return;
        }

        Workbook workbook = null;

        try
        {
            // Load the workbook (no password)
            LoadOptions loadOptions = new LoadOptions();
            workbook = new Workbook(excelPath, loadOptions);
        }
        catch (Exception ex)
        {
            // Detect password protection via exception message
            if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("Error: The workbook is password protected. Provide a password to load it.");
            }
            else
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
            }
            return;
        }

        // Directory where HTML files will be saved
        string outputDir = "HtmlOutput";
        Directory.CreateDirectory(outputDir);

        // Define a common TableCssId to be used for all exported HTML files
        const string commonTableCssId = "myTable";

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            try
            {
                // Set the current worksheet as the active one so that ExportActiveWorksheetOnly works correctly
                workbook.Worksheets.ActiveSheetIndex = sheet.Index;

                // Configure HTML save options
                HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Html)
                {
                    ExportActiveWorksheetOnly = true, // Export only the active worksheet
                    TableCssId = commonTableCssId       // Reuse the same TableCssId for all files
                };

                // Build a safe file name based on the worksheet name
                string safeSheetName = MakeValidFileName(sheet.Name);
                string htmlPath = Path.Combine(outputDir, safeSheetName + ".html");

                // Save the active worksheet as an HTML file
                workbook.Save(htmlPath, options);
                Console.WriteLine($"Worksheet \"{sheet.Name}\" exported to \"{htmlPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error exporting worksheet \"{sheet.Name}\": {ex.Message}");
            }
        }
    }

    // Helper method to replace invalid filename characters
    static string MakeValidFileName(string name)
    {
        foreach (char c in Path.GetInvalidFileNameChars())
        {
            name = name.Replace(c, '_');
        }
        return name;
    }
}
