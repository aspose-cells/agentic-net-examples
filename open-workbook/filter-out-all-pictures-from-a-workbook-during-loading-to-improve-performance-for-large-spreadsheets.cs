using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class FilterOutPicturesOnLoad
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Path to the source workbook (large spreadsheet with pictures)
            string sourcePath = "LargeWorkbookWithPictures.xlsx";

            // Verify source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Create LoadOptions instance with a filter that excludes pictures
            LoadOptions loadOptions = new LoadOptions();
            LoadDataFilterOptions filterOptions = LoadDataFilterOptions.All & ~LoadDataFilterOptions.Picture;
            loadOptions.LoadFilter = new LoadFilter(filterOptions);

            // Load the workbook using the specified options
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Verify that pictures are not loaded (count should be zero)
            int pictureCount = workbook.Worksheets[0].Pictures.Count;
            Console.WriteLine("Number of pictures after loading: " + pictureCount);

            // Save the workbook (optional, demonstrates that saving works without pictures)
            string outputPath = "WorkbookWithoutPictures.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved without pictures to: {outputPath}");
        }
    }
}