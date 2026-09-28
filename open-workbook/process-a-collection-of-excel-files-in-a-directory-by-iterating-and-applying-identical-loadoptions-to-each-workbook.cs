// Title: Load multiple .xlsx workbooks with forced Xlsx LoadOptions, recalculate formulas, and save them to a new folder using Aspose.Cells for .NET
// AI Prompts: Write C# code that enumerates all .xlsx files in a source folder, opens each workbook with LoadOptions set to LoadFormat.Xlsx, invokes CalculateFormula, and writes the workbook to a target directory. | Implement a function that receives input and output paths, applies identical LoadOptions to every Excel file, triggers full formula evaluation, and logs any errors while using Aspose.Cells.
// Common Searches: how to open every .xlsx file in a directory with specific LoadOptions using Aspose.Cells | C# batch evaluate formulas in Excel workbooks with Aspose.Cells | force Xlsx load format for multiple workbooks Aspose.Cells LoadOptions example | save processed Excel files to another folder after calculation in .NET
// Tags: batch load workbooks with LoadOptions Aspose.Cells | evaluate formulas in multiple .xlsx files C# | save processed Excel workbooks to output directory Aspose.Cells | set LoadFormat.Xlsx for workbook loading Aspose.Cells | iterate over Excel files in folder using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program scans a specified input directory for .xlsx files, ensures an output directory exists, and for each file loads the workbook with LoadOptions forcing the Xlsx format, evaluates all formulas, and saves the workbook to the output folder while handling any errors that occur.
class Program
{
    static void Main()
    {
        // Path to the folder containing the source Excel files
        string inputFolder = @"C:\InputExcel";

        // Path to the folder where processed files will be saved
        string outputFolder = @"C:\OutputExcel";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Verify the input directory exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder does not exist: {inputFolder}");
            return;
        }

        // Iterate over all Excel files in the input directory
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
        {
            // Skip if the file somehow does not exist
            if (!File.Exists(filePath))
                continue;

            try
            {
                // Load the workbook with LoadOptions that force XLSX format
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
                Workbook workbook = new Workbook(filePath, loadOptions);

                // Recalculate all formulas
                workbook.CalculateFormula();

                // Build the output file path
                string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));

                // Save the processed workbook (preserving the original format)
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }
    }
}
