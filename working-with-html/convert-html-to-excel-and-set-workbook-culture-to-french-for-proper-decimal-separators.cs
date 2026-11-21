// Title: Convert HTML to Excel (XLSX) and apply French (fr-FR) culture for comma decimal separators using Aspose.Cells for .NET
// Common Searches: how to load an HTML file with Aspose.Cells and save as XLSX using French locale in C# | Aspose.Cells C# convert HTML table to Excel with comma as decimal separator | set workbook culture to fr-FR when converting HTML to Excel with Aspose.Cells | C# Aspose.Cells load HTML using LoadOptions and apply French number formatting | convert HTML to Excel workbook with French culture settings using Aspose.Cells .NET
// Tags: html to xlsx conversion aspose.cells | loadoptions html aspose.cells | apply french culture workbook aspose.cells | comma decimal separator excel c# | save workbook as xlsx aspose.cells

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The example verifies the input HTML file, loads it into an Aspose.Cells Workbook with HTML LoadOptions, sets the workbook's CultureInfo to French (fr-FR) to enforce comma decimal separators, and saves the result as an XLSX file, handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        // Path to the source HTML file
        string htmlFilePath = "input.html";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(htmlFilePath))
        {
            Console.WriteLine($"Error: The file \"{htmlFilePath}\" was not found.");
            return;
        }

        try
        {
            // Set load options for HTML format
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);

            // Load the HTML file into a new workbook instance
            Workbook workbook = new Workbook(htmlFilePath, loadOptions);

            // Set the workbook culture to French (France) to use comma as decimal separator
            workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

            // Save the workbook as an Excel file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors (including Aspose.Cells specific exceptions)
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
