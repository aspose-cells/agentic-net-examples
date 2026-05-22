using System;
using Aspose.Cells;

namespace AsposeCellsAutoFitOnLoad
{
    class Program
    {
        static void Main()
        {
            // Path to the source workbook (replace with actual file path)
            string inputPath = "input.xlsx";

            // Create AutoFitterOptions to enable auto‑fit during loading.
            // Setting OnlyAuto to false ensures all rows are evaluated.
            AutoFitterOptions autoFitOptions = new AutoFitterOptions
            {
                OnlyAuto = false
            };

            // Assign the options to LoadOptions.
            LoadOptions loadOptions = new LoadOptions
            {
                AutoFitterOptions = autoFitOptions
            };

            // Load the workbook with the specified options.
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // At this point row heights have been auto‑fitted automatically.
            // Save the workbook to verify the changes (replace with desired output path).
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook loaded with auto‑fit rows and saved to: {outputPath}");
        }
    }
}