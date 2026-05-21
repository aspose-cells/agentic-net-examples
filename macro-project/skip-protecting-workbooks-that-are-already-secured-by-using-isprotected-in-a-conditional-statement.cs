using System;
using Aspose.Cells;

namespace AsposeCellsProtectionDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook to be processed
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";
            string password = "mySecret";

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Check if the workbook is already protected (structure or window)
            // WorkbookSettings.IsProtected indicates protection status
            if (!workbook.Settings.IsProtected)
            {
                // Workbook is not protected, apply protection with a password
                workbook.Protect(ProtectionType.All, password);
                Console.WriteLine("Workbook was not protected. Protection applied.");
            }
            else
            {
                // Workbook is already protected; skip protecting again
                Console.WriteLine("Workbook is already protected. Skipping protection.");
            }

            // Save the workbook (protected or unchanged)
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}