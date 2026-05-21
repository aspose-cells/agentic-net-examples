using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class ExportToCsvWithoutLoadingModel
{
    static void Main()
    {
        // Path to the existing Excel workbook (can be large)
        string sourceWorkbookPath = "large_input.xlsx";

        // Desired CSV output file path
        string csvOutputPath = "exported_output.csv";

        try
        {
            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourceWorkbookPath))
                throw new FileNotFoundException($"Source workbook not found: {sourceWorkbookPath}");

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(csvOutputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // LoadOptions specify the format of the source file.
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // TxtSaveOptions constructor sets the desired save format (CSV).
            TxtSaveOptions csvSaveOptions = new TxtSaveOptions(SaveFormat.Csv);

            // Perform the conversion in streaming mode without loading the full workbook model.
            ConversionUtility.Convert(sourceWorkbookPath, loadOptions, csvOutputPath, csvSaveOptions);

            Console.WriteLine($"Workbook '{sourceWorkbookPath}' has been exported to CSV at '{csvOutputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during conversion: {ex.Message}");
        }
    }
}