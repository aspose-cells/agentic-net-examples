// Title: Import HTML into Excel with Aspose.Cells and keep <a> tag text as hyperlink display text (C#)
// Common Searches: Aspose.Cells C# import HTML and retain anchor text as hyperlink label | How to preserve <a> tag display text when converting HTML to XLSX using Aspose.Cells | C# code to map HTML hyperlinks to Excel cells with original link text using Aspose.Cells | Load HTML file into Workbook and keep hyperlink text in Excel with Aspose.Cells | Aspose.Cells hyperlink collection after HTML import to set cell value to link text
// Tags: HTML anchor to Excel hyperlink conversion Aspose.Cells | preserve hyperlink display text during HTML import | Workbook Hyperlink collection iteration C# | dynamic casting for Hyperlink row column Aspose.Cells | save workbook as XLSX with hyperlinks Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an HTML file into an Aspose.Cells Workbook, iterates each worksheet's Hyperlink collection, uses dynamic access to retrieve the hyperlink's first row and column, writes the original anchor's display text into the corresponding cell, and saves the result as an XLSX file with functional hyperlinks.
class Program
{
    static void Main()
    {
        const string inputFile = "input.html";
        const string outputFile = "output.xlsx";

        // Verify that the input HTML file exists
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Input file \"{inputFile}\" not found.");
            return;
        }

        try
        {
            // Load the HTML file into a workbook
            var loadOptions = new LoadOptions(LoadFormat.Html);
            var workbook = new Workbook(inputFile, loadOptions);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Process each hyperlink that was created from an <a> tag
                foreach (Hyperlink link in sheet.Hyperlinks)
                {
                    try
                    {
                        // Use dynamic to access FirstRow/FirstColumn (avoids compile‑time binding issues)
                        dynamic dLink = link;
                        int row = dLink.FirstRow;
                        int col = dLink.FirstColumn;

                        // Preserve the anchor's display text as the cell value
                        Cell cell = sheet.Cells[row, col];
                        cell.PutValue(dLink.TextToDisplay);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to process a hyperlink: {ex.Message}");
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputFile);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook as an Excel file
            workbook.Save(outputFile, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
