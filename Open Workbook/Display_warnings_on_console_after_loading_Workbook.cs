using System;
using Aspose.Cells;

namespace AsposeCellsWarningDemo
{
    // Custom warning callback that writes warning details to the console
    public class ConsoleWarningCallback : IWarningCallback
    {
        public void Warning(WarningInfo warningInfo)
        {
            // Display warning type and description
            Console.WriteLine($"Warning Type: {warningInfo.Type}");
            Console.WriteLine($"Description : {warningInfo.Description}");
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Path to the Excel file to be loaded
            string inputPath = "input.xlsx";

            // Create LoadOptions and assign the custom warning callback
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.WarningCallback = new ConsoleWarningCallback();

            // Load the workbook using the options – any warnings during loading will be reported
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Access a cell to demonstrate that the workbook is loaded correctly
            Console.WriteLine($"Cell A1 value: {workbook.Worksheets[0].Cells["A1"].StringValue}");

            // Optionally, save the workbook (warnings during saving can also be captured if needed)
            // workbook.Save("output.xlsx");
        }
    }
}