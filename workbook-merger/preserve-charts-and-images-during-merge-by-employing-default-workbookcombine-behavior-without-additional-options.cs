// Title: Preserve Charts & Images When Merging Excel Workbooks with Aspose.Cells (C#)
// Description: Demonstrates how to combine a source workbook into a destination workbook using Aspose.Cells' Workbook.Combine method. The default combine operation retains all charts, pictures, shapes, and other drawing objects without requiring extra options. The example also shows file‑existence checks and basic error handling, then saves the merged file as a new XLSX document.
// Keywords: Aspose.Cells Workbook.Combine | merge Excel workbooks C# | preserve charts Excel merge | keep images during workbook combine | combine workbooks without losing drawings | Aspose.Cells merge example
// Common Searches: Aspose.Cells merge workbooks keep charts | C# combine Excel files preserving images | Workbook.Combine default behavior | how to merge Excel workbooks without losing drawings | Aspose.Cells combine two workbooks example
// Developer Intent: Combine a source Excel file into a destination file while automatically retaining all embedded charts, pictures, and shapes.
// Use Cases: Create a master report that aggregates monthly sheets without stripping visual analytics. | Merge a template workbook with data‑driven workbooks while preserving branding graphics. | Automate consolidation of departmental spreadsheets, ensuring all inserted diagrams stay intact.
// AI Prompts: Write C# code that uses Aspose.Cells to merge two workbooks and keep every chart and image. | Explain why Workbook.Combine retains drawing objects by default and how to confirm their presence after merging. | Suggest robust error‑handling patterns for workbook merging when files may be missing or corrupted.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to combine a source workbook into a destination workbook using Aspose.Cells' Workbook.Combine method. The default combine operation retains all charts, pictures, shapes, and other drawing objects without requiring extra options. The example also shows file‑existence checks and basic error handling, then saves the merged file as a new XLSX document.
    public class PreserveChartsAndImagesDuringMerge
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            const string destPath = "DestinationWorkbook.xlsx";
            const string sourcePath = "SourceWorkbook.xlsx";
            const string outputPath = "CombinedWorkbook.xlsx";

            try
            {
                // Ensure destination workbook exists; create an empty one if missing
                if (!File.Exists(destPath))
                {
                    new Workbook().Save(destPath);
                }

                // Ensure source workbook exists; create an empty one if missing
                if (!File.Exists(sourcePath))
                {
                    new Workbook().Save(sourcePath);
                }

                // Load the destination workbook (the workbook that will receive the merged content)
                Workbook destWorkbook = new Workbook(destPath);

                // Load the source workbook (the workbook whose charts and images will be merged)
                Workbook sourceWorkbook = new Workbook(sourcePath);

                // Combine the source workbook into the destination workbook.
                // This operation preserves charts, images, shapes, and other drawing objects.
                destWorkbook.Combine(sourceWorkbook);

                // Save the combined workbook.
                destWorkbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine("Workbooks combined successfully. Charts and images are preserved.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
