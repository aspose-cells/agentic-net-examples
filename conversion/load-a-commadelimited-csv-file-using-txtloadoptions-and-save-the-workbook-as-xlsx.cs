// Title: C# – Load a comma‑separated CSV using Aspose.Cells TxtLoadOptions and save as XLSX
// AI Prompts: Write C# code that reads a CSV file with TxtLoadOptions (separator ',', ConvertNumericData = true, ConvertDateTimeData = true) and saves it as an XLSX workbook using Aspose.Cells. | Create a method that verifies a CSV file exists, loads it with Aspose.Cells TxtLoadOptions, enables numeric/date conversion, and exports the workbook to XLSX. | Generate error‑handling logic for a CSV‑to‑XLSX conversion in C#, catching missing‑file and other exceptions while using Aspose.Cells.
// Common Searches: aspnet convert csv file to xlsx using txtloadoptions aspocells c# | how to set comma separator in Aspose.Cells TxtLoadOptions when loading CSV | c# Aspose.Cells load csv with numeric and date conversion and save as xlsx | example code for converting comma delimited csv to excel workbook with Aspose.Cells | handling file not found exception during csv to xlsx conversion aspocells
// Tags: Aspose.Cells TxtLoadOptions CSV loading | CSV to XLSX conversion C# Aspose.Cells | comma separator configuration Aspose.Cells | numeric and datetime conversion Aspose.Cells | file existence check C# Aspose.Cells conversion

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example checks for a comma‑delimited CSV file, loads it into an Aspose.Cells Workbook with TxtLoadOptions (comma separator, numeric and date conversion enabled), and saves the workbook as an XLSX file, handling missing‑file and other exceptions.
    public class CsvToXlsxConverter
    {
        public static void Run()
        {
            // Path to the source CSV file (comma‑delimited)
            string csvPath = "input.csv";

            // Path for the resulting XLSX file
            string xlsxPath = "output.xlsx";

            // Verify that the CSV file exists to avoid FileNotFoundException
            if (!File.Exists(csvPath))
            {
                Console.WriteLine($"CSV file not found: {csvPath}");
                return;
            }

            try
            {
                // Create TxtLoadOptions for loading a CSV file
                TxtLoadOptions loadOptions = new TxtLoadOptions
                {
                    // Explicitly set the separator to comma (default is also comma)
                    Separator = ',',
                    // Enable conversion of numeric and date values
                    ConvertNumericData = true,
                    ConvertDateTimeData = true
                };

                // Load the CSV file into a workbook using the specified load options
                Workbook workbook = new Workbook(csvPath, loadOptions);

                // Save the workbook as XLSX
                workbook.Save(xlsxPath, SaveFormat.Xlsx);

                Console.WriteLine($"CSV file \"{csvPath}\" has been successfully converted to \"{xlsxPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during conversion: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
