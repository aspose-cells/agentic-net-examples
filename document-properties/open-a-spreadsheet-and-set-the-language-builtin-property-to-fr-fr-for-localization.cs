// Title: How to set the built‑in Language property to French (fr-FR) in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that opens an existing .xlsx file (or creates a new workbook), sets the built‑in Language document property to "fr-FR", and saves the workbook. | Show a .NET console example that modifies the Language built‑in property of an Excel file for localization using Aspose.Cells.
// Common Searches: Aspose.Cells C# set workbook language to fr-FR | Change Excel file locale programmatically using Aspose.Cells .NET | Set built‑in document property Language in .xlsx with C# | How to localize an Excel workbook by updating Language property via Aspose.Cells | C# example for modifying built‑in Language property of an Excel workbook
// Tags: Aspose.Cells set built-in Language property | C# modify Excel document properties | localize Excel workbook Aspose.Cells | set workbook locale fr-FR .NET | update built-in properties in .xlsx with Aspose

using Aspose.Cells;
using System;
using System.IO;

// Loads an existing Excel file (or creates a new workbook), assigns the built‑in Language property the value "fr-FR" for French (France) localization, and saves the workbook as output.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            Workbook workbook;

            // Load the existing spreadsheet if it exists; otherwise create a new workbook
            const string inputPath = "input.xlsx";
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
            }

            // Set the built‑in Language property to French (France)
            var languageProp = workbook.BuiltInDocumentProperties["Language"];
            languageProp.Value = "fr-FR";

            // Save the modified spreadsheet
            const string outputPath = "output.xlsx";
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            // Log or handle exceptions as needed
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
