using System;
using Aspose.Cells;

namespace AsposeCellsPasswordLoadDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the password‑protected workbook
            string filePath = "protected.xlsx";

            // Prompt the user to enter the password
            Console.Write("Enter password to open the workbook: ");
            string password = Console.ReadLine();

            // Create LoadOptions and set the entered password
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = password;

            // Load the workbook using the load options
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Example: read the value of cell A1 from the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Console.WriteLine("Value of A1: " + sheet.Cells["A1"].Value?.ToString());

            // Keep console open
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
    }
}