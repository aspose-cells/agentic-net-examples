// Title: Merge Multiple .xls Workbooks into a Single .xlsx Using Aspose.Cells for .NET (C#)
// Description: C# sample that loads each .xls file with new Workbook(filePath), combines them into a destination workbook via Workbook.Combine, skips missing files, logs progress, and saves the merged result as MergedResult.xlsx with full error handling.
// Keywords: Aspose.Cells | C# | merge Excel workbooks | combine .xls files | Workbook.Combine | load workbook from file | save as .xlsx | batch Excel processing | error handling | .NET Excel automation
// Common Searches: C# merge multiple xls files Aspose.Cells | Combine .xls workbooks into one .xlsx .NET | Aspose.Cells Workbook.Combine example | Load Excel file with new Workbook(filePath) C# | Skip missing Excel files during merge Aspose
// Developer Intent: Programmatically combine several legacy .xls workbooks into a single .xlsx file using Aspose.Cells for .NET.
// Use Cases: Consolidate daily sales .xls reports into a master workbook for quarterly analysis. | Migrate legacy .xls archives to .xlsx format in an automated batch job. | Create a scheduled task that merges incoming .xls data files while handling missing or corrupted files.
// AI Prompts: Write C# code that reads a list of .xls paths, merges them with Aspose.Cells Workbook.Combine, and saves as .xlsx with robust error handling. | Show how to modify the merge loop to include only worksheets whose names match a specific pattern. | Generate a PowerShell script that invokes a compiled .NET assembly to merge .xls files using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// C# sample that loads each .xls file with new Workbook(filePath), combines them into a destination workbook via Workbook.Combine, skips missing files, logs progress, and saves the merged result as MergedResult.xlsx with full error handling.
class Program
{
    static void Main()
    {
        // Paths of source workbooks to merge.
        string[] sourceFiles = { "Source1.xls", "Source2.xls", "Source3.xls" };

        // Destination workbook that will contain merged data.
        Workbook destWorkbook = new Workbook();

        foreach (string filePath in sourceFiles)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}. Skipping.");
                    continue;
                }

                // Load source workbook and combine it with the destination.
                Workbook srcWorkbook = new Workbook(filePath);
                destWorkbook.Combine(srcWorkbook);
                Console.WriteLine($"Merged: {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing {filePath}: {ex.Message}");
            }
        }

        try
        {
            // Save the merged workbook.
            destWorkbook.Save("MergedResult.xlsx", SaveFormat.Xlsx);
            Console.WriteLine("Merged workbook saved as MergedResult.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving merged workbook: {ex.Message}");
        }
    }
}
