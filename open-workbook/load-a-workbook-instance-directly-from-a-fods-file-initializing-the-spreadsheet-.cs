using System;
using Aspose.Cells;

namespace AsposeCellsFodsLoadDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the FODS file
            string fodsPath = "sample.fods";

            // Create LoadOptions specifying the FODS format
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Fods);

            // Load the workbook from the FODS file using the constructor that accepts a file path and LoadOptions
            Workbook workbook = new Workbook(fodsPath, loadOptions);

            // The workbook is now initialized and ready for further processing.
            // Example: display the name of the first worksheet and the value of cell A1.
            Worksheet firstSheet = workbook.Worksheets[0];
            Console.WriteLine("First worksheet name: " + firstSheet.Name);
            Console.WriteLine("Cell A1 value: " + firstSheet.Cells["A1"].StringValue);
        }
    }
}