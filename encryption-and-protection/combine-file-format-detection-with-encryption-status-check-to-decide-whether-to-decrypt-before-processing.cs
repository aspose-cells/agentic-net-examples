using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file
            string sourceFilePath = "input.xlsx";

            // Password to use if the file is encrypted (replace with actual password)
            string password = "yourPassword";

            // Detect the file format and encryption status
            FileFormatInfo fileInfo = FileFormatUtil.DetectFileFormat(sourceFilePath);
            Console.WriteLine($"Detected format: {fileInfo.FileFormatType}");
            Console.WriteLine($"Is encrypted: {fileInfo.IsEncrypted}");

            Workbook workbook;

            if (fileInfo.IsEncrypted)
            {
                // File is encrypted – load it with the provided password
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
                loadOptions.Password = password;
                workbook = new Workbook(sourceFilePath, loadOptions);
                Console.WriteLine("Workbook loaded with password.");
            }
            else
            {
                // File is not encrypted – load normally
                workbook = new Workbook(sourceFilePath);
                Console.WriteLine("Workbook loaded without password.");
            }

            // Example processing: write a value into cell A1 of the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Processed");

            // Save the processed workbook to a new file
            string outputFilePath = "output.xlsx";
            workbook.Save(outputFilePath);
            Console.WriteLine($"Processed workbook saved to: {outputFilePath}");
        }
    }
}