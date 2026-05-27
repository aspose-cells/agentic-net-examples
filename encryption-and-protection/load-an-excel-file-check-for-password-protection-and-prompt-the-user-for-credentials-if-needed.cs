using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPasswordDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the Excel file
            string filePath = "sample.xlsx";

            // Verify that the file exists before attempting to load it
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                // Detect file format and check if the workbook is encrypted
                FileFormatInfo fileInfo = FileFormatUtil.DetectFileFormat(filePath);
                Workbook workbook;

                if (fileInfo.IsEncrypted)
                {
                    // Prompt the user for the password
                    Console.Write("The workbook is password protected. Enter password: ");
                    string password = Console.ReadLine();

                    // Load the workbook using the provided password
                    LoadOptions loadOptions = new LoadOptions { Password = password };
                    workbook = new Workbook(filePath, loadOptions);
                }
                else
                {
                    // Load the workbook directly (no password required)
                    workbook = new Workbook(filePath);
                }

                // Example operation: display the value of cell A1 from the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Console.WriteLine("Value of A1: " + sheet.Cells["A1"].Value?.ToString());

                // (Optional) Save the workbook after any modifications
                // workbook.Save("output.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}