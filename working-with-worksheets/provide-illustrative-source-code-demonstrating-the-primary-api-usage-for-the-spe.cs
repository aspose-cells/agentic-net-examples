using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsLicenseDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Initialize the License object
            var license = new License();

            // Attempt to set the license if the file exists
            string licensePath = "Aspose.Cells.NET.lic";
            if (File.Exists(licensePath))
            {
                license.SetLicense(licensePath);
            }
            else
            {
                Console.WriteLine($"License file '{licensePath}' not found. Continuing without a license.");
            }

            // Create a new workbook (uses the licensed component if available)
            var workbook = new Workbook();

            // Access the first worksheet
            var worksheet = workbook.Worksheets[0];

            // Write a sample value to cell A1
            worksheet.Cells["A1"].PutValue("Hello, Aspose.Cells with License!");

            // Save the workbook to verify that the license is applied correctly
            string outputPath = "LicensedWorkbook.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
    }
}