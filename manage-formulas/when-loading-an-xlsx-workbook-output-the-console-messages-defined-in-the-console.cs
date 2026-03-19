using System;
using Aspose.Cells;

namespace AsposeCellsLoadDemo
{
    // Custom warning callback that writes warning details to the console
    public class ConsoleWarningCallback : IWarningCallback
    {
        public void Warning(WarningInfo warningInfo)
        {
            // Output warning type and description
            Console.WriteLine($"Warning: {warningInfo.WarningType} - {warningInfo.Description}");
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file to be loaded
            string filePath = "sample.xlsx";

            // Initialize load options and assign the custom warning callback
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.WarningCallback = new ConsoleWarningCallback();

            // Load the workbook using the constructor that accepts file path and LoadOptions
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Output basic information after loading
            Console.WriteLine($"Workbook loaded successfully.");
            Console.WriteLine($"Number of worksheets: {workbook.Worksheets.Count}");

            // Example access to a cell (optional, may trigger additional warnings)
            if (workbook.Worksheets.Count > 0)
            {
                Worksheet sheet = workbook.Worksheets[0];
                Console.WriteLine($"First worksheet name: {sheet.Name}");
                Console.WriteLine($"Cell A1 value: {sheet.Cells["A1"].StringValue}");
            }
        }
    }
}