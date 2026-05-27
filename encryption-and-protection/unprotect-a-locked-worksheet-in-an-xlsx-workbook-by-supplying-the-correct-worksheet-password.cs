using System;
using System.IO;
using Aspose.Cells;

namespace WorksheetUnprotectDemo
{
    class Program
    {
        static void Main()
        {
            const string inputPath = "protected.xlsx";
            const string outputPath = "unprotected.xlsx";
            const string password = "myPassword";

            try
            {
                // Ensure the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Unprotect the worksheet using the password
                worksheet.Unprotect(password);

                // Verify that the worksheet is no longer protected
                Console.WriteLine($"Worksheet is protected: {worksheet.IsProtected}");

                // Save the unprotected workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any errors (including invalid password)
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}