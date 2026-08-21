// Title: Load UTF-16 CSV with Asian characters using Aspose.Cells TxtLoadOptions (C#)
// Description: Shows how to create a TxtLoadOptions object, set Encoding = Encoding.Unicode, load a UTF‑16 CSV that contains Japanese, Korean or Chinese text, read a cell to verify Unicode handling, and save the workbook as XLSX.
// Keywords: Aspose.Cells | TxtLoadOptions | UTF-16 | Unicode CSV | C# | .NET | Asian characters | Japanese CSV | Korean CSV | Chinese CSV | CSV to XLSX conversion | encoding property
// Common Searches: Aspose.Cells load UTF-16 CSV | C# read Unicode CSV with Aspose | Import Asian characters CSV .NET | Convert UTF-16 CSV to Excel using Aspose.Cells | TxtLoadOptions encoding example
// Developer Intent: Import a UTF‑16 encoded CSV that contains Asian Unicode text into an Aspose.Cells Workbook and optionally export it to another format.
// Use Cases: Preserve Japanese, Korean, or Chinese characters when importing CSV data. | Validate that a specific cell (e.g., A1) contains the expected Unicode string after load. | Batch‑convert UTF‑16 CSV files to XLSX for reporting pipelines. | Handle missing‑file scenarios gracefully during CSV import. | Integrate Unicode CSV ingestion into automated .NET data workflows.
// AI Prompts: Generate C# code that uses Aspose.Cells TxtLoadOptions to load a UTF‑16 CSV with Japanese text and saves it as XLSX. | Explain how setting TxtLoadOptions.Encoding to Encoding.Unicode ensures correct import of Chinese characters from a CSV file. | Provide a step‑by‑step example for reading a UTF‑16 CSV containing Korean characters, checking cell A1, and handling FileNotFoundException.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to create a TxtLoadOptions object, set Encoding = Encoding.Unicode, load a UTF‑16 CSV that contains Japanese, Korean or Chinese text, read a cell to verify Unicode handling, and save the workbook as XLSX.
    public class LoadCsvUtf16Demo
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            // Path to the CSV file that is saved with UTF‑16 (Unicode) encoding
            string csvPath = "asian_utf16.csv";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(csvPath))
            {
                Console.WriteLine($"Error: The file \"{csvPath}\" was not found.");
                return;
            }

            try
            {
                // Create load options for text files and set the encoding to UTF‑16
                TxtLoadOptions loadOptions = new TxtLoadOptions
                {
                    Encoding = Encoding.Unicode // UTF‑16 Little Endian
                };

                // Load the CSV file using the specified options
                Workbook workbook = new Workbook(csvPath, loadOptions);

                // Access the first worksheet and display a cell that contains Asian characters
                Worksheet sheet = workbook.Worksheets[0];
                Console.WriteLine("A1 value: " + sheet.Cells["A1"].StringValue);

                // Save the workbook to another format (e.g., XLSX) to verify the import succeeded
                string outputPath = "converted.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook successfully saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the CSV file:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
