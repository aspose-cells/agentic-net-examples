// Title: Custom IFilePathProvider for Aspose.Cells HTML Export – Save Worksheets to a Separate Folder with Full‑Path Links (C#)
// Description: Demonstrates how to implement IFilePathProvider to place each worksheet’s HTML file in a dedicated "Sheets" directory, replace spaces with underscores, and keep navigation intact by enabling IsFullPathLink. The workbook is saved as a main HTML page with linked sheet files.
// Keywords: Aspose.Cells IFilePathProvider | HTML export custom folder | C# Aspose.Cells HtmlSaveOptions | Save worksheets as separate HTML files | IsFullPathLink example | Aspose.Cells custom file naming | Aspose.Cells multi‑sheet HTML export
// Common Searches: Aspose.Cells custom IFilePathProvider C# | export workbook to HTML each sheet in its own folder | how to keep links working when saving Aspose.Cells to HTML | Aspose.Cells HtmlSaveOptions FilePathProvider usage | C# save Excel sheets as separate HTML files
// Developer Intent: Create a safe, folder‑based naming scheme for worksheet HTML files and preserve hyperlink integrity during Aspose.Cells HTML export.
// Use Cases: Organize multi‑sheet HTML output into a structured directory for web publishing. | Apply corporate naming conventions (underscores, specific folders) to exported files. | Maintain functional navigation between the main HTML page and individual sheet pages.
// AI Prompts: Generate C# code that implements IFilePathProvider to store each worksheet as "Sheets/SheetName.htm" and integrates it with HtmlSaveOptions. | Show how to export an Aspose.Cells workbook to HTML with custom file paths and full‑path links, handling spaces in sheet names safely. | Explain the role of IsFullPathLink when using a custom IFilePathProvider to keep cross‑sheet hyperlinks operational.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to implement IFilePathProvider to place each worksheet’s HTML file in a dedicated "Sheets" directory, replace spaces with underscores, and keep navigation intact by enabling IsFullPathLink. The workbook is saved as a main HTML page with linked sheet files.
public class CustomFilePathProvider : IFilePathProvider
{
    // Returns the full path for a given worksheet name.
    public string GetFullName(string sheetName)
    {
        const string folder = "Sheets";
        Directory.CreateDirectory(folder); // Ensure the folder exists.

        // Replace spaces with underscores to create a safe file name.
        string safeName = sheetName.Replace(' ', '_');

        // Build the full path (e.g., Sheets/Sheet1.htm).
        return Path.Combine(folder, $"{safeName}.htm");
    }
}

public class ExportWorkbookToHtml
{
    public static void Run()
    {
        try
        {
            // Create a workbook with two worksheets and add sample data.
            Workbook workbook = new Workbook();
            workbook.Worksheets[0].Name = "First Sheet";
            workbook.Worksheets[0].Cells["A1"].PutValue("Data in first sheet");

            workbook.Worksheets.Add("Second Sheet");
            workbook.Worksheets[1].Cells["B2"].PutValue("Data in second sheet");

            // Configure HTML save options.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportActiveWorksheetOnly = false,          // Export all worksheets.
                FilePathProvider = new CustomFilePathProvider(), // Custom file naming.
                IsFullPathLink = true                       // Use full path links.
            };

            // Save the workbook. Main file: "Workbook.html", worksheets in "Sheets" folder.
            workbook.Save("Workbook.html", saveOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during export: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main()
    {
        ExportWorkbookToHtml.Run();
    }
}
