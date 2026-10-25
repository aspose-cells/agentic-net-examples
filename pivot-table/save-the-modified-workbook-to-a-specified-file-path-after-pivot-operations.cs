// Title: Refresh all pivot tables and save the modified Excel workbook to a custom path using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells refresh all pivot tables before saving the workbook | How to save an Excel file to a specific folder after pivot table updates with Aspose.Cells | Aspose.Cells ensure output directory exists when calling Workbook.Save in .NET | Programmatically refresh pivot tables in an existing workbook using Aspose.Cells | Save modified workbook to a dynamic file path using Aspose.Cells example
// Tags: Aspose.Cells pivot table refresh | Workbook.Save custom location C# | create output directory before saving Excel Aspose.Cells | programmatic pivot table refresh .NET | save modified Excel workbook Aspose.Cells example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable class

namespace AsposeCellsExample
{
    // The PivotWorkbookHandler class loads an existing Excel file, iterates through each worksheet to refresh every pivot table, ensures the target output directory exists, and then saves the modified workbook to the specified path using Aspose.Cells for .NET.
    public class PivotWorkbookHandler
    {
        private readonly Workbook _workbook;

        // Load the workbook (uses the load rule)
        public PivotWorkbookHandler(string sourcePath)
        {
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException($"Source file not found: {sourcePath}");

            _workbook = new Workbook(sourcePath);
        }

        // Perform pivot table modifications here
        public void ModifyPivotTables()
        {
            try
            {
                // Refresh all pivot tables in the workbook
                foreach (Worksheet sheet in _workbook.Worksheets)
                {
                    foreach (PivotTable pt in sheet.PivotTables)
                    {
                        pt.RefreshData();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error while modifying pivot tables: {ex.Message}");
                throw;
            }
        }

        // Save the modified workbook to the specified file path (uses the save rule)
        public void Save(string outputPath)
        {
            // Ensure the output directory exists
            string directory = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                Directory.CreateDirectory(directory);

            _workbook.Save(outputPath);
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Example input and output file paths; adjust as needed
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                var handler = new PivotWorkbookHandler(inputPath);
                handler.ModifyPivotTables(); // Apply modifications
                handler.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (FileNotFoundException ex)
            {
                Console.Error.WriteLine(ex.Message);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
