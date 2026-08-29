// Title: Merge two Excel workbooks and save the combined file to a custom path using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells to combine a source Workbook into a destination Workbook and save the result to a specified file path in C#. | Create two Excel workbooks, add data, merge them with Workbook.Combine, then call Workbook.Save with a custom filename. | Demonstrate persisting a merged workbook to a user‑defined location using destWorkbook.Save in Aspose.Cells for .NET.
// Common Searches: how to merge two Excel files and save as new file using Aspose.Cells in C# | Aspose.Cells combine workbooks then save to specific folder .NET | C# code example for Workbook.Combine and Workbook.Save with Aspose.Cells | save merged workbook to custom path Aspose.Cells .NET example | combine source workbook into destination workbook and export as xlsx using Aspose.Cells
// Tags: Aspose.Cells Workbook.Combine usage C# | Aspose.Cells Workbook.Save with explicit file path | merge Excel files into single workbook Aspose.Cells | define output file path for merged workbook | .NET example of combining and persisting workbooks

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The sample creates a source and a destination Workbook, writes sample data to each, merges the source into the destination using Workbook.Combine, and then saves the combined workbook to a user‑specified file path (CombinedWorkbook.xlsx) with Workbook.Save.
    public class MergeAndSaveDemo
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create the first workbook (source)
            Workbook sourceWorkbook = new Workbook();
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            sourceSheet.Name = "SourceSheet";
            sourceSheet.Cells["A1"].PutValue("Data from source workbook");

            // Create the second workbook (destination)
            Workbook destWorkbook = new Workbook(FileFormatType.Xlsx);
            Worksheet destSheet = destWorkbook.Worksheets[0];
            destSheet.Name = "DestinationSheet";
            destSheet.Cells["B2"].PutValue("Data from destination workbook");

            // Combine the source workbook into the destination workbook
            destWorkbook.Combine(sourceWorkbook);

            // Define the output file path (extension determines format)
            string outputPath = "CombinedWorkbook.xlsx";

            // Save the combined workbook
            destWorkbook.Save(outputPath);

            Console.WriteLine($"Combined workbook saved to: {outputPath}");
        }
    }
}
