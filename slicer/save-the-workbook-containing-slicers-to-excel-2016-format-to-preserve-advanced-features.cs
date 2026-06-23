using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSlicerSaveDemo
{
    class Program
    {
        static void Main()
        {
            const string inputFile = "WorkbookWithSlicers.xlsx";
            const string outputFile = "WorkbookWithSlicers_Saved.xlsx";

            try
            {
                // Verify that the source workbook exists before attempting to load it.
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Error: Input file '{inputFile}' was not found.");
                    return;
                }

                // Load the existing workbook that contains slicers.
                Workbook workbook = new Workbook(inputFile);

                // Save the workbook in Excel 2016 (XLSX) format to preserve slicers and other advanced features.
                workbook.Save(outputFile, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}