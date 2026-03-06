using System;
using System.IO;
using System.Text;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Determine the full path to the CSV file
        string csvFileName = "special_chars.csv";
        string csvPath = Path.Combine(Environment.CurrentDirectory, csvFileName);

        // If the CSV file does not exist, create a sample one with Unicode characters
        if (!File.Exists(csvPath))
        {
            string[] lines =
            {
                "\"Name\",\"Comment\"",
                "\"Alice\",\"Hello, world!\"",
                "\"Bob\",\"Привет мир\"",
                "\"Chloé\",\"¡Hola, mundo!\""
            };
            File.WriteAllLines(csvPath, lines, Encoding.UTF8);
        }

        // Configure load options for CSV handling
        TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv)
        {
            Separator = ',',
            HasTextQualifier = true,
            TextQualifier = '"',
            Encoding = Encoding.UTF8,
            IsMultiEncoded = true
        };

        // Load the CSV file into a workbook using the configured options
        Workbook workbook = new Workbook(csvPath, loadOptions);

        // Save the workbook as an Excel file
        string outputPath = Path.Combine(Environment.CurrentDirectory, "output.xlsx");
        workbook.Save(outputPath, SaveFormat.Xlsx);

        Console.WriteLine($"CSV file loaded and saved as Excel: {outputPath}");
    }
}