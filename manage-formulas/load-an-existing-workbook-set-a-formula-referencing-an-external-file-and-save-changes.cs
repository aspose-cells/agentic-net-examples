// Title: Add an external workbook reference formula, refresh linked data, and save with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to load an existing workbook, create or register an external link to another XLSX file, assign a formula that points to a cell in the external workbook, refresh the linked data source, and save the updated file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells external link | C# set formula from another workbook | UpdateLinkedDataSource Aspose.Cells | refresh linked data Excel | save workbook with external references | programmatic external workbook reference | .NET Excel automation
// Common Searches: Aspose.Cells add external reference C# | How to refresh linked data source in Aspose.Cells | Set formula to another workbook using Aspose.Cells | Save Excel file with external links .NET | Create external link programmatically Aspose.Cells
// Developer Intent: Programmatically link a cell to an external workbook, recalculate the link, and persist the changes.
// Use Cases: Consolidate monthly reports into a single summary workbook. | Build a dashboard that pulls key metrics from multiple source files. | Automate financial models that depend on data stored in separate spreadsheets.
// AI Prompts: Generate C# code with Aspose.Cells to add an external link and set a formula referencing a cell in the linked file. | Explain the role of UpdateLinkedDataSource when working with external references in Aspose.Cells. | Provide robust error‑handling patterns for missing external workbooks while setting formulas in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to load an existing workbook, create or register an external link to another XLSX file, assign a formula that points to a cell in the external workbook, refresh the linked data source, and save the updated file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Define file names
            string mainPath = "Main.xlsx";
            string externalPath = "External.xlsx";

            // Ensure the main workbook exists; create a simple one if missing
            if (!File.Exists(mainPath))
            {
                Workbook wb = new Workbook();
                wb.Worksheets[0].Name = "Sheet1";
                wb.Save(mainPath);
            }

            // Ensure the external workbook exists; create a simple one with sample data if missing
            if (!File.Exists(externalPath))
            {
                Workbook wb = new Workbook();
                Worksheet ws = wb.Worksheets[0];
                ws.Name = "Sheet1";
                ws.Cells["A2"].PutValue(123); // sample value for the external reference
                wb.Save(externalPath);
            }

            // Load the existing main workbook
            Workbook mainWorkbook = new Workbook(mainPath);
            // Set the FileName property – required when the workbook is opened from a stream and contains external references.
            mainWorkbook.FileName = mainPath;

            // Register the external link (if not already present)
            string[] externalSheets = new string[] { "Sheet1" };
            int externalLinkIndex = mainWorkbook.Worksheets.ExternalLinks.Add(externalPath, externalSheets);

            // Set a formula that references a cell in the external workbook.
            Worksheet mainWs = mainWorkbook.Worksheets[0];
            mainWs.Cells["A1"].Formula = $"=[{Path.GetFileName(externalPath)}]Sheet1!$A$2";

            // Load the external workbook so that linked data can be refreshed.
            Workbook externalWorkbook = new Workbook(externalPath);

            // Update the external data source – this makes the formula calculate correctly.
            mainWorkbook.UpdateLinkedDataSource(new Workbook[] { externalWorkbook });

            // Save the modified workbook.
            string outputPath = "Main_Updated.xlsx";
            mainWorkbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
