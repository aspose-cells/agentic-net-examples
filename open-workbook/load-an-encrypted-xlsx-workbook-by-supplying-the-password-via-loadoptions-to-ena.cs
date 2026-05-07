using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class LoadEncryptedWorkbookDemo
    {
        public static void Run()
        {
            // Path to the encrypted workbook
            string filePath = "protected.xlsx";

            // Create LoadOptions and set the password required to open the workbook
            LoadOptions loadOptions = new LoadOptions
            {
                Password = "test"
            };

            // Load the encrypted workbook using the password
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Example: read a cell value to verify successful loading
            string cellValue = workbook.Worksheets[0].Cells["A1"].StringValue;
            Console.WriteLine("Cell A1 value: " + cellValue);

            // Remove password protection and save an unprotected copy (optional)
            workbook.Settings.Password = null;
            workbook.Save("unprotected.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            LoadEncryptedWorkbookDemo.Run();
        }
    }
}