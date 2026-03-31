using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPasswordValidation
{
    class Program
    {
        static void Main(string[] args)
        {
            // Expect two arguments: file path and password
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: AsposeCellsPasswordValidation <filePath> <password>");
                return;
            }

            string filePath = args[0];
            string password = args[1];

            // Verify the password without opening the workbook
            bool isPasswordValid;
            using (FileStream stream = File.OpenRead(filePath))
            {
                // FileFormatUtil.VerifyPassword works for encrypted OOXML (XLSX) and ODS files
                isPasswordValid = FileFormatUtil.VerifyPassword(stream, password);
            }

            if (!isPasswordValid)
            {
                Console.WriteLine("Password validation failed. Cannot open the workbook.");
                return;
            }

            Console.WriteLine("Password is valid. Loading workbook...");

            // Load the workbook using the verified password
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = password;
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Confirm that the workbook is no longer considered encrypted after loading
            Console.WriteLine($"Workbook IsEncrypted after load: {workbook.Settings.IsEncrypted}");

            // Example processing: output the value of the first cell
            var firstCellValue = workbook.Worksheets[0].Cells[0, 0].Value;
            Console.WriteLine($"First cell value: {firstCellValue}");
        }
    }
}