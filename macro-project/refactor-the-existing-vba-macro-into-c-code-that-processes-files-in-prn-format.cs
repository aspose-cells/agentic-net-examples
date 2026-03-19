using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPrnProcessor
{
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing PRN files
            string inputFolder = @"C:\InputPrnFiles";
            // Folder where converted XLSX files will be saved
            string outputFolder = @"C:\OutputXlsxFiles";

            // Verify input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            // Ensure output directory exists
            Directory.CreateDirectory(outputFolder);

            // Process each .prn file in the input folder
            foreach (string prnPath in Directory.GetFiles(inputFolder, "*.prn"))
            {
                try
                {
                    // Load the PRN file using automatic format detection
                    var loadOptions = new LoadOptions(LoadFormat.Auto);
                    var workbook = new Workbook(prnPath, loadOptions);

                    // If the workbook contains VBA/macros, remove them
                    if (workbook.HasMacro)
                    {
                        workbook.RemoveMacro();
                    }

                    // Build the output file name with .xlsx extension
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(prnPath);
                    string xlsxPath = Path.Combine(outputFolder, fileNameWithoutExt + ".xlsx");

                    // Save the workbook as XLSX (macro‑free)
                    workbook.Save(xlsxPath, SaveFormat.Xlsx);

                    Console.WriteLine($"Successfully converted '{prnPath}' to '{xlsxPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{prnPath}': {ex.Message}");
                }
            }

            Console.WriteLine("Processing completed.");
        }
    }
}