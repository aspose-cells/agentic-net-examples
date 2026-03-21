using System;
using Aspose.Cells;

namespace AsposeCellsLoadValidateDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file (adjust as needed)
            string sourcePath = "example.xlsx";

            // Load the workbook using the string constructor
            Workbook workbook = new Workbook(sourcePath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Display basic information to validate successful load
            Console.WriteLine("Worksheet Name: " + sheet.Name);
            Console.WriteLine("Number of Cells: " + sheet.Cells.Count);

            // Save the workbook in MHTML format to demonstrate saving works
            string outputPath = "example_output.mht";
            workbook.Save(outputPath, SaveFormat.MHtml);

            // Clean up
            workbook.Dispose();

            Console.WriteLine("Workbook loaded and saved successfully.");
        }
    }
}