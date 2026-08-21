// Title: Load Workbook via LightCells, Keep Defined Names Containing "Total", Save as PDF (C# Aspose.Cells)
// Description: Demonstrates how to use Aspose.Cells LightCells API with a LoadFilter that loads only defined‑name objects, retain names that include the word "Total", discard the rest, and export the filtered workbook to PDF in C#.
// Keywords: Aspose.Cells LightCells API | LoadFilter DefinedNames | C# filter defined names | export Excel to PDF | keep named ranges Total | remove unwanted defined names | LoadOptions LoadFilter | Aspose.Cells PDF conversion
// Common Searches: Aspose.Cells load only defined names C# | filter named ranges containing Total Aspose.Cells | LightCells API export to PDF | remove Excel defined names before PDF conversion | C# load workbook with LoadFilter.DefinedNames
// Developer Intent: Load an Excel file with only defined‑name objects, keep those whose name contains "Total", delete the others, and save the result as a PDF.
// Use Cases: Generate a concise PDF report that shows only total‑related named ranges from a large workbook. | Automate cleanup of named ranges in a CI/CD pipeline before publishing Excel files as PDFs. | Create lightweight PDFs for dashboards by stripping unrelated defined names.
// AI Prompts: Write C# code using Aspose.Cells to load an Excel workbook with LoadFilter.DefinedNames, keep defined names that contain a specific keyword, and save the workbook as a PDF. | Show how to remove unwanted defined names after loading a workbook with LightCells API and then export the filtered workbook to PDF. | Explain the configuration of LoadOptions and LoadFilter to load only defined names and filter them by text in Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to use Aspose.Cells LightCells API with a LoadFilter that loads only defined‑name objects, retain names that include the word "Total", discard the rest, and export the filtered workbook to PDF in C#.
class LoadFilterDefinedNamesToPdf
{
    static void Main()
    {
        // Paths for source Excel file and destination PDF file
        string sourcePath = "input.xlsx";
        string pdfPath = "output.pdf";

        // Verify that the source file exists to avoid FileNotFoundException
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Error: Source file not found at '{sourcePath}'.");
            return;
        }

        try
        {
            // Create LoadOptions and set a LoadFilter that loads only defined names
            LoadOptions loadOptions = new LoadOptions
            {
                // Load only the defined name objects (no cell data, charts, etc.)
                LoadFilter = new LoadFilter(LoadDataFilterOptions.DefinedNames)
            };

            // Load the workbook using LightCells API (LoadOptions)
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Retrieve all defined names in the workbook
            Name[] allNames = workbook.Worksheets.Names.Filter(NameScopeType.All, -1);

            // Collect names that need to be removed
            var namesToRemove = new System.Collections.Generic.List<string>();

            // Iterate through the names and keep only those that contain "Total"
            foreach (Name name in allNames)
            {
                // 'Text' holds the defined name identifier
                if (!string.IsNullOrEmpty(name.Text) &&
                    name.Text.IndexOf("Total", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Console.WriteLine($"Keeping defined name: {name.Text} -> {name.RefersTo}");
                }
                else
                {
                    // Mark names that do not contain "Total" for removal
                    namesToRemove.Add(name.Text);
                }
            }

            // Remove the unwanted defined names
            foreach (string nameToRemove in namesToRemove)
            {
                workbook.Worksheets.Names.Remove(nameToRemove);
            }

            // Save the workbook as PDF
            workbook.Save(pdfPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved to PDF: {pdfPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
