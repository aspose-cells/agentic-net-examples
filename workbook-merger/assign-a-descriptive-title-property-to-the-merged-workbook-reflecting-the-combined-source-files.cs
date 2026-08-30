// Title: Assign a descriptive document Title to a workbook after combining multiple Excel files with Aspose.Cells in C#
// AI Prompts: Generate C# code that merges several .xlsx files using Aspose.Cells and sets the workbook's BuiltInDocumentProperties.Title to a comma‑separated list of the source filenames. | Show how to update the Title property of an Aspose.Cells Workbook after calling the Combine method. | Explain the steps to programmatically assign a custom document title to a merged Excel workbook in a .NET application.
// Common Searches: C# Aspose.Cells set workbook title after combining multiple spreadsheets | How to update BuiltInDocumentProperties.Title for a merged workbook using Aspose.Cells .NET | Aspose.Cells combine Excel files and assign custom document title programmatically | Set document properties of merged workbook with Aspose.Cells in C#
// Tags: aspose.cells combine workbooks title property | c# set builtindocumentproperties title aspose.cells | merge multiple xlsx files aspose.cells | custom document title for merged workbook .net

using System;
using System.IO;
using System.Linq;
using Aspose.Cells;

namespace AsposeCellsMergeDemo
{
    // The example loads three source Excel files, merges them into a single Workbook using Aspose.Cells' Combine method, assigns a descriptive Title built‑in document property that lists the source file names, and saves the merged workbook as MergedWorkbook.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Define source Excel files to be merged
                string[] sourceFiles = new string[]
                {
                    "Source1.xlsx",
                    "Source2.xlsx",
                    "Source3.xlsx"
                };

                // Define the output merged workbook file name
                string outputFile = "MergedWorkbook.xlsx";

                // Create an empty workbook that will hold the merged content
                Workbook mergedWorkbook = new Workbook();

                // Iterate through each source file, load it, and combine its content into the merged workbook
                foreach (string filePath in sourceFiles)
                {
                    // Verify the source file exists to avoid FileNotFoundException
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"Source file not found and will be skipped: {filePath}");
                        continue;
                    }

                    // Load the source workbook from file
                    Workbook sourceWorkbook = new Workbook(filePath);

                    // Combine the source workbook into the merged workbook
                    mergedWorkbook.Combine(sourceWorkbook);
                }

                // Assign a descriptive title reflecting the combined source files
                mergedWorkbook.BuiltInDocumentProperties.Title =
                    "Combined Workbook: " + string.Join(", ", sourceFiles.Select(f => Path.GetFileNameWithoutExtension(f)));

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputFile));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the merged workbook to the specified output file
                mergedWorkbook.Save(outputFile, SaveFormat.Xlsx);
                Console.WriteLine($"Merged workbook saved successfully to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
