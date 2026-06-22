using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class SetPrintTitleColumnsDemo
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set columns A and B to repeat on the left side of each printed page
            worksheet.PageSetup.PrintTitleColumns = "$A:$B";

            // Add sample data to illustrate the repeated columns (optional)
            for (int i = 0; i < 30; i++)
            {
                worksheet.Cells[$"A{i + 1}"].PutValue($"Header A {i + 1}");
                worksheet.Cells[$"B{i + 1}"].PutValue($"Header B {i + 1}");
                worksheet.Cells[$"C{i + 1}"].PutValue($"Data {i + 1}");
            }

            // Determine output file path
            string outputFile = "PrintTitleColumnsAB.xlsx";

            // Save the workbook
            workbook.Save(outputFile);
        }
    }
}