using System;
using Aspose.Cells;

namespace AsposeCellsWorksheetUnprotectDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load the protected workbook (replace with your actual file path)
            string inputPath = "ProtectedWorkbook.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Access the worksheet you want to unprotect (e.g., the first worksheet)
            Worksheet worksheet = workbook.Worksheets[0];

            // Unprotect the worksheet using the password it was protected with
            // If the worksheet was protected without a password, use worksheet.Unprotect();
            string password = "yourPassword"; // replace with the actual password
            worksheet.Unprotect(password);

            // Save the unprotected workbook to a new file
            string outputPath = "UnprotectedWorkbook.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Worksheet '{worksheet.Name}' has been unprotected and saved to '{outputPath}'.");
        }
    }
}