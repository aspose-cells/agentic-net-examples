using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Author: Aspose.Cells .NET example – loading a password‑protected XLSX using LoadOptions
    class Program
    {
        static void Main()
        {
            // Create load options and set the decryption password
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = "myPassword";

            // Load the protected workbook using the load options
            Workbook workbook = new Workbook("protected.xlsx", loadOptions);

            // Example of using LightCells API (optional) – here we just access a cell value
            // LightCellsDataProvider provider = new LightCellsDataProvider(workbook);
            // LightCellsProcessor processor = new LightCellsProcessor();
            // processor.Process(provider);

            // Output a cell value to verify successful loading
            Console.WriteLine("Cell A1 value: " + workbook.Worksheets[0].Cells["A1"].Value);
        }
    }
}