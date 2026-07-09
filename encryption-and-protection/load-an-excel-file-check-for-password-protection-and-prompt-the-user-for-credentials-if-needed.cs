using System;
using Aspose.Cells;

namespace AsposeCellsPasswordDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file
            string filePath = "sample.xlsx";

            // Detect file format and check if the workbook is encrypted
            FileFormatInfo fileInfo = FileFormatUtil.DetectFileFormat(filePath);
            Workbook workbook;

            if (fileInfo.IsEncrypted)
            {
                // Prompt the user for the password
                Console.Write("The workbook is password protected. Please enter the password: ");
                string password = Console.ReadLine();

                // Load the workbook using the provided password
                LoadOptions loadOptions = new LoadOptions();
                loadOptions.Password = password;

                try
                {
                    workbook = new Workbook(filePath, loadOptions);
                    Console.WriteLine("Workbook loaded successfully with the provided password.");
                }
                catch (Exception ex)
                {
                    // Handle incorrect password or other loading errors
                    Console.WriteLine($"Failed to load workbook: {ex.Message}");
                    return;
                }
            }
            else
            {
                // Load the workbook normally (no password required)
                workbook = new Workbook(filePath);
                Console.WriteLine("Workbook loaded successfully (no password required).");
            }

            // Example: display the value of the first cell in the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Console.WriteLine($"First cell value: {sheet.Cells["A1"].Value}");
        }
    }
}