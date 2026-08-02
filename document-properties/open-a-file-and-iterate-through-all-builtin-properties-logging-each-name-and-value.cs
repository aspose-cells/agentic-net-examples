using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsExamples
{
    class BuiltInPropertiesReader
    {
        static void Main()
        {
            // Path to the Excel file to be opened
            string filePath = "input.xlsx";

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(filePath);

            // Retrieve the collection of built‑in document properties
            BuiltInDocumentPropertyCollection builtInProps = workbook.BuiltInDocumentProperties;

            // Iterate through each property and log its name and value
            foreach (DocumentProperty prop in builtInProps)
            {
                Console.WriteLine($"{prop.Name}: {prop.Value}");
            }
        }
    }
}