// Title: How to remove external links and data connections from an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an Excel file with Aspose.Cells, deletes every DataConnection object, clears any formula that references another workbook, and saves the result to a new file. | Implement a reusable C# method that accepts input and output paths, strips external links (both data connections and external formulas) from the workbook using Aspose.Cells, and returns a success status.
// Common Searches: asp.net remove data connections from Excel using aspose.cells | c# clear formulas that reference other workbooks with aspose.cells | how to make an Excel file self-contained with aspose.cells in .net | strip external links from workbook programmatically aspose.cells | delete external data connections and formulas in Excel via C#
// Tags: Aspose.Cells purge data connections | Aspose.Cells purge external workbook formulas | Aspose.Cells save workbook without external links | C# strip external references from Excel workbook | Aspose.Cells workbook cleanup external references

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing workbook, removes all DataConnection objects, iterates through each worksheet and cell to clear formulas that contain external workbook references, ensures the output directory exists, and saves the modified workbook as a self‑contained file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Clear any external data connections
            try
            {
                if (workbook.DataConnections != null && workbook.DataConnections.Count > 0)
                {
                    workbook.DataConnections.Clear();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Unable to clear data connections: {ex.Message}");
            }

            // Remove external references that may still exist inside formulas
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Cell cell in sheet.Cells)
                {
                    try
                    {
                        if (cell.IsFormula && !string.IsNullOrEmpty(cell.Formula) && cell.Formula.Contains("["))
                        {
                            cell.Formula = string.Empty;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Unable to process cell {cell.Name} on sheet \"{sheet.Name}\": {ex.Message}");
                    }
                }
            }

            // Ensure the output directory exists
            try
            {
                string? outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Unable to create output directory: {ex.Message}");
            }

            // Save the workbook as a self‑contained file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
