using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsMhtJsonToCsv
{
    public class Converter
    {
        public static void Run(string mhtFilePath, string csvOutputPath)
        {
            if (!File.Exists(mhtFilePath))
                throw new FileNotFoundException($"MHT file not found: {mhtFilePath}");

            // Load the MHT file content as text
            string mhtContent = File.ReadAllText(mhtFilePath, Encoding.UTF8);

            // Extract JSON part
            string jsonPattern = @"(\{[\s\S]*?\})|(\[[\s\S]*?\])";
            Match match = Regex.Match(mhtContent, jsonPattern);
            if (!match.Success)
                throw new InvalidDataException("No JSON content found in the MHT file.");

            string json = match.Value.Trim();

            // Create workbook and import JSON
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            JsonLayoutOptions jsonOptions = new JsonLayoutOptions
            {
                ArrayAsTable = true,
                ConvertNumericOrDate = true,
                DateFormat = "yyyy-MM-dd",
                NumberFormat = "0.##"
            };

            JsonUtility.ImportData(json, cells, 0, 0, jsonOptions);

            // Save as CSV
            workbook.Save(csvOutputPath, SaveFormat.Csv);
        }
    }

    class Program
    {
        static void Main()
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string mhtPath = Path.Combine(baseDir, "input.mht");
            string csvPath = Path.Combine(baseDir, "output.csv");

            // Create a sample MHT file with embedded JSON if it does not exist
            if (!File.Exists(mhtPath))
            {
                string sampleJson = "[{\"Name\":\"John\",\"Age\":30,\"JoinDate\":\"2023-01-15\"},{\"Name\":\"Jane\",\"Age\":25,\"JoinDate\":\"2022-07-08\"}]";
                string mhtTemplate = $"Content-Type: multipart/related; boundary=\"----=_NextPart_000_0000\"\r\n\r\n------=_NextPart_000_0000\r\nContent-Type: application/json\r\n\r\n{sampleJson}\r\n------=_NextPart_000_0000--";
                File.WriteAllText(mhtPath, mhtTemplate, Encoding.UTF8);
            }

            Converter.Run(mhtPath, csvPath);

            Console.WriteLine($"Conversion completed successfully. CSV saved to: {csvPath}");
        }
    }
}