using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ConvertTextToNumberAndExport
    {
        public static void Run()
        {
            // Create a new workbook (in-memory)
            Workbook workbook = new Workbook();

            // Access the first worksheet and its cells collection
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Sample CSV data containing numeric values as text and some non‑numeric strings
            string csvData = "ID,Value,Date,Note\n" +
                             "1,\"123.45\",\"2023-01-15\",\"Valid\"\n" +
                             "2,\"678\",\"2023-02-20\",\"Valid\"\n" +
                             "3,\"ABC\",\"2023-03-10\",\"Invalid\"";

            // Convert the CSV string to a memory stream (UTF‑8 encoding)
            using (MemoryStream csvStream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(csvData)))
            {
                // Import the CSV data into the worksheet.
                // Parameters: stream, delimiter, convertNumericData, firstRow, firstColumn
                // Setting convertNumericData to true enables automatic conversion of numeric strings.
                cells.ImportCSV(csvStream, ",", true, 0, 0);
            }

            // Save the workbook as XLSX
            workbook.Save("ConvertedData.xlsx", SaveFormat.Xlsx);

            // Optional: display the converted values to verify
            Console.WriteLine("Converted values:");
            Console.WriteLine($"A2 (ID)    : {cells["A2"].IntValue} (Type: {cells["A2"].Type})");
            Console.WriteLine($"B2 (Value) : {cells["B2"].DoubleValue} (Type: {cells["B2"].Type})");
            Console.WriteLine($"C2 (Date)  : {cells["C2"].DateTimeValue} (Type: {cells["C2"].Type})");
            Console.WriteLine($"D2 (Note)  : {cells["D2"].StringValue} (Type: {cells["D2"].Type})");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ConvertTextToNumberAndExport.Run();
        }
    }
}