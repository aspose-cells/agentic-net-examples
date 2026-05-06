using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class TxtToJsonConverter
    {
        public static void Run()
        {
            // Path to the source TXT file (Excel‑compatible text format)
            string txtFilePath = "source.txt";

            // Ensure the source file exists; create a sample if it does not.
            if (!File.Exists(txtFilePath))
            {
                // Sample tab‑delimited data
                string[] sampleLines =
                {
                    "Name\tAge\tJoinDate",
                    "Alice\t30\t2022-01-15",
                    "Bob\t25\t2023-03-10",
                    "Charlie\t35\t2021-07-22"
                };
                File.WriteAllLines(txtFilePath, sampleLines);
            }

            // Load the TXT file into a workbook.
            TxtLoadOptions loadOptions = new TxtLoadOptions
            {
                Separator = '\t',
                ConvertNumericData = true,
                ConvertDateTimeData = true
            };
            Workbook workbook = new Workbook(txtFilePath, loadOptions);

            // Configure JSON save options (default exports data only)
            JsonSaveOptions jsonOptions = new JsonSaveOptions();

            // Destination JSON file
            string jsonOutputPath = "output.json";

            // Save the workbook as JSON using the configured options
            workbook.Save(jsonOutputPath, jsonOptions);

            // Optional: display the generated JSON content
            Console.WriteLine("JSON export completed. Content:");
            Console.WriteLine(File.ReadAllText(jsonOutputPath));
        }

        // Entry point
        public static void Main(string[] args)
        {
            Run();
        }
    }
}