using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCsvExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source CSV file
            string csvPath = "input.csv";

            // Ensure the CSV file exists (for demonstration purposes we create a simple one)
            if (!File.Exists(csvPath))
            {
                File.WriteAllText(csvPath, "Name,Age,Note\nJohn,30,\"Invalid#Char\"\nAlice,25,\"Another?Char\"");
            }

            // Create CSV load options and set the separator
            TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv);
            loadOptions.Separator = ',';

            // Load the CSV file into a workbook
            Workbook workbook = new Workbook(csvPath, loadOptions);

            // Replace invalid characters in the worksheet data.
            workbook.Replace("#", string.Empty);
            workbook.Replace("?", string.Empty);

            // Save the resulting workbook as an Excel file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"CSV file '{csvPath}' has been loaded, invalid characters replaced, and saved as '{outputPath}'.");
        }
    }
}