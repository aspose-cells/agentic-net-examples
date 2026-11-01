// Title: Load multiple .xls workbooks into Aspose.Cells Workbook objects with file‑existence checks and error handling in C#
// AI Prompts: Generate C# code that iterates over a list of .xls file paths, verifies each file exists, and creates an Aspose.Cells Workbook instance for every valid file while catching and logging load exceptions. | Write a method that accepts a collection of Excel file paths, loads each into a List<Workbook> using Aspose.Cells, skips missing files, and returns the successfully loaded workbooks with detailed error messages.
// Common Searches: c# load a list of .xls files into Aspose.Cells Workbook objects and skip missing files | how to safely create Aspose.Cells Workbook instances from multiple Excel file paths in C# | example code for loading multiple Excel workbooks with Aspose.Cells and handling load exceptions | validate file existence before loading workbooks using Aspose.Cells in a C# application
// Tags: load xls workbooks Aspose.Cells C# | file existence validation before Aspose.Cells workbook creation | exception handling Aspose.Cells workbook loading | batch loading Excel workbooks with Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace WorkbookMergeExample
{
    // The sample defines a list of source .xls file paths, checks each path for existence, loads each file into an Aspose.Cells Workbook via the Workbook constructor, adds the workbook to a collection, logs successful loads or errors, and finally reports the total number of workbooks loaded.
    class Program
    {
        static void Main(string[] args)
        {
            // List of source XLS file paths to be merged
            List<string> sourceFilePaths = new List<string>
            {
                @"C:\Data\Source1.xls",
                @"C:\Data\Source2.xls",
                @"C:\Data\Source3.xls"
            };

            // Collection to hold loaded workbooks
            List<Workbook> sourceWorkbooks = new List<Workbook>();

            // Load each source workbook safely
            foreach (string filePath in sourceFilePaths)
            {
                try
                {
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found: {filePath}. Skipping.");
                        continue;
                    }

                    Workbook wb = new Workbook(filePath); // Load workbook
                    sourceWorkbooks.Add(wb);
                    Console.WriteLine($"Loaded workbook: {Path.GetFileName(filePath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error loading workbook '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine($"Loaded {sourceWorkbooks.Count} workbooks successfully.");
        }
    }
}
