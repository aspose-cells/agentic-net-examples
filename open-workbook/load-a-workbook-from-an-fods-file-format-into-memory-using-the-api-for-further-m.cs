using System;
using Aspose.Cells;

namespace AsposeCellsFodsLoadExample
{
    class Program
    {
        static void Main()
        {
            // Path to the FODS file to be loaded
            string fodsPath = "sample.fods";

            // Create LoadOptions specifying the FODS format
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Fods);

            // Load the workbook from the FODS file using the constructor with LoadOptions
            Workbook workbook = new Workbook(fodsPath, loadOptions);

            // Example manipulation: display the number of worksheets loaded
            Console.WriteLine($"Workbook loaded. Worksheet count: {workbook.Worksheets.Count}");

            // Example manipulation: access the first worksheet and read a cell value
            Worksheet firstSheet = workbook.Worksheets[0];
            Console.WriteLine($"First sheet name: {firstSheet.Name}");
            Console.WriteLine($"Cell A1 value: {firstSheet.Cells["A1"].StringValue}");
        }
    }
}