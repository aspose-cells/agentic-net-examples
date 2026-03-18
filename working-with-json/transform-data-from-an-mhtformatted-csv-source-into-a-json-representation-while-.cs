using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Utility;

class MhtCsvToJsonConverter
{
    static void Main()
    {
        // Path to the MHT file that contains the CSV data
        string mhtFilePath = "source.mht";

        // If the MHT file does not exist, create a simple one for demonstration purposes
        if (!File.Exists(mhtFilePath))
        {
            string sampleCsv = "Name,Age,City\r\nAlice,30,New York\r\nBob,25,London\r\n";
            string mhtTemplate = "From: <Saved by WebKit>\r\nSubject: CSV Export\r\nMIME-Version: 1.0\r\nContent-Type: multipart/related; boundary=\"----=_NextPart_000_0000\"\r\n\r\n------=_NextPart_000_0000\r\nContent-Type: text/plain; charset=\"utf-8\"\r\nContent-Transfer-Encoding: 8bit\r\n\r\n" + sampleCsv + "\r\n------=_NextPart_000_0000--";
            File.WriteAllText(mhtFilePath, mhtTemplate, Encoding.UTF8);
        }

        // Read the entire MHT content
        string mhtContent = File.ReadAllText(mhtFilePath, Encoding.UTF8);

        // Simple extraction: CSV part is after the first empty line following the MIME headers
        string[] parts = mhtContent.Split(new[] { "\r\n\r\n", "\n\n" }, StringSplitOptions.None);
        string csvData = parts.Length > 1 ? parts[1] : string.Empty;

        if (string.IsNullOrWhiteSpace(csvData))
        {
            Console.WriteLine("No CSV data found in the MHT file.");
            return;
        }

        // Convert CSV string to a memory stream (UTF‑8 encoded)
        byte[] csvBytes = Encoding.UTF8.GetBytes(csvData);
        using (MemoryStream csvStream = new MemoryStream(csvBytes))
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Import the CSV data into the worksheet starting at cell A1 (row 0, column 0)
            cells.ImportCSV(csvStream, ",", true, 0, 0);

            // Determine the used range (including header row)
            int maxRow = cells.MaxDataRow;
            int maxColumn = cells.MaxDataColumn;
            Aspose.Cells.Range usedRange = cells.CreateRange(0, 0, maxRow + 1, maxColumn + 1);

            // Configure JSON export options
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ExportEmptyCells = true,
                HasHeaderRow = true,
                ExportNestedStructure = false
            };

            // Export the range to a JSON string
            string jsonResult = JsonUtility.ExportRangeToJson(usedRange, jsonOptions);

            // Output the JSON (or write to a file)
            Console.WriteLine(jsonResult);
            File.WriteAllText("output.json", jsonResult, Encoding.UTF8);
        }
    }
}