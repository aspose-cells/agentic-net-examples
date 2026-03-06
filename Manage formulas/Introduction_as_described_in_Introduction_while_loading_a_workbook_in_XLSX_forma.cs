using System;
using Aspose.Cells;

namespace AsposeCellsIntroduction
{
    // Demonstrates loading an existing XLSX workbook, accessing its content,
    // and saving it to a new file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file.
            string sourcePath = "input.xlsx";

            // Load the workbook using the constructor that accepts a file name.
            Workbook workbook = new Workbook(sourcePath);

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Display worksheet information.
            Console.WriteLine("Worksheet Name: " + sheet.Name);
            Console.WriteLine("Cell A1 Value: " + sheet.Cells["A1"].StringValue);

            // Save the workbook to a new file.
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine("Workbook loaded from '" + sourcePath + "' and saved as '" + outputPath + "'.");
        }
    }
}