// Title: C# – Load an Excel workbook without pictures using Aspose.Cells LoadFilter
// Description: Demonstrates how to create a LoadOptions object with a LoadFilter that excludes pictures (LoadDataFilterOptions.All & ~LoadDataFilterOptions.Picture), load a large workbook efficiently, verify picture collections are empty, and optionally save the picture‑free file.
// Keywords: Aspose.Cells LoadFilter picture exclusion | C# load Excel without images | skip pictures Aspose.Cells | improve Excel load performance | LoadDataFilterOptions picture | large workbook memory optimization | Aspose.Cells GitHub example
// Common Searches: Aspose.Cells load workbook without pictures | exclude images when opening Excel in .NET | speed up loading large Excel files Aspose | C# LoadFilter to ignore pictures | remove pictures during workbook load
// Developer Intent: Load an Excel file while omitting all embedded pictures to reduce memory consumption and accelerate processing.
// Use Cases: Data‑only analysis of massive spreadsheets where graphics are irrelevant. | Batch conversion of Excel files to other formats without carrying over images. | Generating lightweight reports from source workbooks that contain many pictures.
// AI Prompts: Show C# code to open an Excel workbook with Aspose.Cells while skipping all pictures. | Explain how to use LoadDataFilterOptions to exclude pictures and load only values and formulas. | Provide a method to confirm that no pictures were loaded after applying a LoadFilter.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPictureFilterDemo
{
    // Demonstrates how to create a LoadOptions object with a LoadFilter that excludes pictures (LoadDataFilterOptions.All & ~LoadDataFilterOptions.Picture), load a large workbook efficiently, verify picture collections are empty, and optionally save the picture‑free file.
    class Program
    {
        static void Main()
        {
            // Path to the source workbook (large spreadsheet with many pictures)
            string sourcePath = "LargeWorkbookWithPictures.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            try
            {
                // Create LoadOptions instance
                LoadOptions loadOptions = new LoadOptions();

                // Configure a LoadFilter to load everything except pictures
                // LoadDataFilterOptions.All includes all data; we remove the Picture flag
                LoadDataFilterOptions filterOptions = LoadDataFilterOptions.All & ~LoadDataFilterOptions.Picture;
                loadOptions.LoadFilter = new LoadFilter(filterOptions);

                // Load the workbook with the specified options
                Workbook workbook = new Workbook(sourcePath, loadOptions);

                // Verify that pictures are not loaded
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Console.WriteLine($"Worksheet '{sheet.Name}' picture count: {sheet.Pictures.Count}");
                }

                // Save the workbook (optional, to a new file)
                string outputPath = "WorkbookWithoutPictures.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine("Workbook loaded without pictures and saved to: " + outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
