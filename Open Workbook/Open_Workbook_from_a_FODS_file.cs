using System;
using Aspose.Cells;

namespace AsposeCellsFodsExample
{
    class Program
    {
        static void Main()
        {
            // Specify the path to the FODS file
            string fodsPath = "sample.fods";

            // Create LoadOptions with the FODS load format
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Fods);

            // Open the workbook using the constructor that accepts a file path and LoadOptions
            using (Workbook workbook = new Workbook(fodsPath, loadOptions))
            {
                // Output the name of the first worksheet to verify successful loading
                Console.WriteLine("First worksheet name: " + workbook.Worksheets[0].Name);
            }
        }
    }
}