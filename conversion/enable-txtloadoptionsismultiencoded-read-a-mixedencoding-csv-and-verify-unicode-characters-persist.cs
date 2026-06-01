using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsMixedEncodingDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Register code page provider for Windows‑1252 (required on .NET Core/5+)
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

                // Path for the temporary CSV file
                string csvPath = Path.Combine(Path.GetTempPath(), "mixed_encoding.csv");

                // Prepare two lines with different encodings
                string utf8Line = "名前,年齢\n";               // Japanese characters (UTF‑8)
                string win1252Line = "John,30\n";            // ASCII / Windows‑1252

                // Write the first line using UTF‑8 encoding and the second using Windows‑1252
                using (FileStream fs = new FileStream(csvPath, FileMode.Create, FileAccess.Write))
                {
                    byte[] utf8Bytes = Encoding.UTF8.GetBytes(utf8Line);
                    fs.Write(utf8Bytes, 0, utf8Bytes.Length);

                    byte[] win1252Bytes = Encoding.GetEncoding(1252).GetBytes(win1252Line);
                    fs.Write(win1252Bytes, 0, win1252Bytes.Length);
                }

                // Ensure the CSV file exists before loading
                if (!File.Exists(csvPath))
                    throw new FileNotFoundException("The CSV file was not created.", csvPath);

                // Configure load options to handle multiple encodings
                TxtLoadOptions loadOptions = new TxtLoadOptions
                {
                    IsMultiEncoded = true,          // Enable mixed‑encoding support
                    Encoding = Encoding.UTF8        // Default encoding (used for the first part)
                };

                // Load the CSV file with the specified options
                Workbook workbook = new Workbook(csvPath, loadOptions);
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Verify that Unicode characters are preserved
                Console.WriteLine("Cell A1 (Japanese Name): " + cells["A1"].StringValue);
                Console.WriteLine("Cell B1 (Japanese Age Header): " + cells["B1"].StringValue);
                Console.WriteLine("Cell A2 (English Name): " + cells["A2"].StringValue);
                Console.WriteLine("Cell B2 (English Age): " + cells["B2"].StringValue);

                // Optional: save to an Excel file to confirm successful import
                string outputPath = Path.Combine(Path.GetTempPath(), "mixed_encoding_output.xlsx");
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
            finally
            {
                // Clean up temporary CSV file
                string csvPath = Path.Combine(Path.GetTempPath(), "mixed_encoding.csv");
                if (File.Exists(csvPath))
                {
                    try
                    {
                        File.Delete(csvPath);
                    }
                    catch (Exception delEx)
                    {
                        Console.WriteLine("Failed to delete temporary CSV file: " + delEx.Message);
                    }
                }
            }
        }
    }
}