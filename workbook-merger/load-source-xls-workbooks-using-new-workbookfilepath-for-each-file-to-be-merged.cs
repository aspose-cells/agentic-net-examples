// Title: How to load multiple .xls workbooks into a List<Workbook> using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells batch load .xls files into a workbook collection | Aspose.Cells load several Excel files with file‑existence check in .NET | Create a List<Workbook> from multiple Excel file paths using Aspose.Cells | Handle exceptions while opening multiple .xls workbooks with Aspose.Cells in C# | Iterate over file paths to instantiate Aspose.Cells Workbook objects
// Tags: load multiple xls workbooks Aspose.Cells C# | Aspose.Cells workbook collection from file paths | file existence validation before Workbook creation | exception handling for Excel workbook loading | batch processing of Excel files Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example shows how to define a list of .xls file paths, verify each file's existence, instantiate an Aspose.Cells Workbook for each existing file, add the workbooks to a List<Workbook>, and log successes or errors, finally reporting the total number of loaded workbooks.
class Program
{
    static void Main()
    {
        // List of source Excel file paths to be merged
        List<string> sourceFiles = new List<string>
        {
            @"C:\Data\File1.xls",
            @"C:\Data\File2.xls",
            @"C:\Data\File3.xls"
        };

        // Collection to hold loaded workbooks
        List<Workbook> workbooks = new List<Workbook>();

        // Load each workbook safely
        foreach (string filePath in sourceFiles)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}. Skipping.");
                    continue;
                }

                Workbook wb = new Workbook(filePath);
                workbooks.Add(wb);
                Console.WriteLine($"Loaded workbook: {Path.GetFileName(filePath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading workbook '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine($"Total workbooks loaded: {workbooks.Count}");
        // Further processing can be performed on the 'workbooks' collection.
    }
}
