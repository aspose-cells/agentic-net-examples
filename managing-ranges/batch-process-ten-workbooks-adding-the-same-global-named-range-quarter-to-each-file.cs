// Title: C# – Batch add a workbook‑level named range “Quarter” to multiple Excel files with Aspose.Cells
// Description: Iterates through ten Excel workbooks, loads each with Aspose.Cells, creates the global named range “Quarter” (refers to Sheet1!$A$1:$B$4) only if it does not exist, optionally sorts the name collection, and saves the modified files to a “Processed” folder while handling missing files and runtime errors.
// Keywords: Aspose.Cells | C# | .NET | global named range | workbook‑level name | batch process Excel files | NameCollection | add named range programmatically | multiple workbooks automation | Excel error handling
// Common Searches: add a workbook level named range to many Excel files using Aspose.Cells | C# batch create global named range in multiple workbooks | Aspose.Cells loop through files to add same named range | how to check and add a named range only if missing in Excel with Aspose | sort name collection before saving workbook Aspose.Cells
// Developer Intent: Create the global named range "Quarter" in each of ten Excel workbooks, adding it only when absent.
// Use Cases: Standardize a quarterly data reference across all monthly report workbooks before consolidation. | Prepare a set of template spreadsheets for a financial model that requires a common named range for downstream calculations. | Automate the migration of legacy Excel files to include a required global named range for a new reporting engine.
// AI Prompts: Write C# code with Aspose.Cells that adds a global named range to every Excel file in a folder, skipping files that already contain the name. | Explain how to verify the existence of a workbook‑level named range across multiple workbooks and add it if missing, including best‑practice error handling. | Provide a script that processes a list of Excel files, creates a named range, sorts the NameCollection, and saves the results to a separate directory.

using System;
using System.IO;
using Aspose.Cells;

// Iterates through ten Excel workbooks, loads each with Aspose.Cells, creates the global named range “Quarter” (refers to Sheet1!$A$1:$B$4) only if it does not exist, optionally sorts the name collection, and saves the modified files to a “Processed” folder while handling missing files and runtime errors.
class BatchAddGlobalNamedRange
{
    static void Main()
    {
        // Paths of the ten workbooks to process
        string[] inputFiles = new string[10]
        {
            "file1.xlsx",
            "file2.xlsx",
            "file3.xlsx",
            "file4.xlsx",
            "file5.xlsx",
            "file6.xlsx",
            "file7.xlsx",
            "file8.xlsx",
            "file9.xlsx",
            "file10.xlsx"
        };

        // Folder where processed workbooks will be saved
        string outputFolder = "Processed";
        Directory.CreateDirectory(outputFolder);

        foreach (string inputPath in inputFiles)
        {
            try
            {
                // Verify that the source file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}. Skipping.");
                    continue;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the global name collection
                NameCollection names = workbook.Worksheets.Names;

                // Add the global named range "Quarter" if it does not already exist
                if (names["Quarter"] == null)
                {
                    int index = names.Add("Quarter");          // Define a new name
                    Name quarterName = names[index];
                    quarterName.RefersTo = "=Sheet1!$A$1:$B$4"; // Example reference
                    quarterName.SheetIndex = 0;                // 0 = global (workbook‑level) scope
                }

                // Optional: sort names for better performance before saving
                workbook.Worksheets.SortNames();

                // Save the modified workbook to the output folder
                string fileName = Path.GetFileName(inputPath);
                string outputPath = Path.Combine(outputFolder, fileName);
                workbook.Save(outputPath);
                Console.WriteLine($"Processed and saved: {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors and continue with the next file
                Console.WriteLine($"Error processing '{inputPath}': {ex.Message}");
            }
        }
    }
}
