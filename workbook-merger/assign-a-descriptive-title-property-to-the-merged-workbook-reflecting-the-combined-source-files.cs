// Title: Set Workbook Title After Merging Excel Files with Aspose.Cells for .NET
// Description: Shows how to load two Excel workbooks, combine them into a new workbook using Aspose.Cells, assign a descriptive BuiltInDocumentProperties.Title that lists the source filenames, and save the merged file.
// Keywords: Aspose.Cells | C# merge workbooks | set workbook title | BuiltInDocumentProperties | Excel combine | document metadata | merged workbook | .NET Excel library
// Common Searches: Aspose.Cells set title after combining workbooks | C# merge multiple Excel files and update document title | How to change BuiltInDocumentProperties.Title in a merged workbook | Assign custom title to combined Excel workbook using Aspose.Cells | Update Excel workbook metadata after merging files
// Developer Intent: Add or modify the Title built‑in document property of a workbook created by merging several Excel files with Aspose.Cells.
// Use Cases: Consolidate monthly financial sheets into a single report and embed the month names in the Title for quick reference. | Create a master data workbook from departmental files and record the source filenames in the Title for auditability. | Automate the merging of template workbooks and set a Title that reflects the combined template versions.
// AI Prompts: Generate C# code that accepts a list of Excel file paths, merges them with Aspose.Cells, and sets the BuiltInDocumentProperties.Title to a comma‑separated list of the source names. | Provide a method to merge an arbitrary number of workbooks and update the Title property to include the total count and each filename. | Explain how to read, modify, and save the Title property of a workbook after using Aspose.Cells to combine other workbooks.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to load two Excel workbooks, combine them into a new workbook using Aspose.Cells, assign a descriptive BuiltInDocumentProperties.Title that lists the source filenames, and save the merged file.
    public class MergeWorkbooksWithTitleDemo
    {
        public static void Run()
        {
            try
            {
                // Paths of source workbooks to be merged
                string sourcePath1 = "Source1.xlsx";
                string sourcePath2 = "Source2.xlsx";

                // Ensure source files exist; create empty workbooks if missing
                if (!File.Exists(sourcePath1))
                {
                    new Workbook().Save(sourcePath1, SaveFormat.Xlsx);
                }
                if (!File.Exists(sourcePath2))
                {
                    new Workbook().Save(sourcePath2, SaveFormat.Xlsx);
                }

                // Load the source workbooks
                Workbook sourceWorkbook1 = new Workbook(sourcePath1);
                Workbook sourceWorkbook2 = new Workbook(sourcePath2);

                // Create a destination workbook (empty)
                Workbook mergedWorkbook = new Workbook();

                // Combine the source workbooks into the destination
                mergedWorkbook.Combine(sourceWorkbook1);
                mergedWorkbook.Combine(sourceWorkbook2);

                // Set a descriptive title reflecting the combined files
                mergedWorkbook.BuiltInDocumentProperties.Title =
                    $"Combined Workbook: {Path.GetFileName(sourcePath1)}, {Path.GetFileName(sourcePath2)}";

                // Save the merged workbook
                string outputPath = "MergedWorkbook.xlsx";
                mergedWorkbook.Save(outputPath, SaveFormat.Xlsx);

                // Confirmation
                Console.WriteLine($"Workbooks merged and saved to '{outputPath}' with title: {mergedWorkbook.BuiltInDocumentProperties.Title}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during merging: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            MergeWorkbooksWithTitleDemo.Run();
        }
    }
}
