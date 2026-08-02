using System;
using Aspose.Cells;

namespace AsposeCellsCommandLine
{
    class Program
    {
        // Usage:
        //   AsposeCellsCommandLine.exe <excelFilePath> [propertyName propertyValue]
        // If propertyName and propertyValue are supplied, the property is added to the workbook.
        static void Main(string[] args)
        {
            if (args.Length < 1)
            {
                Console.WriteLine("Error: No file path supplied.");
                Console.WriteLine("Usage: AsposeCellsCommandLine.exe <excelFilePath> [propertyName propertyValue]");
                return;
            }

            string filePath = args[0];

            // Load the workbook using the constructor that accepts a file path.
            using (Workbook workbook = new Workbook(filePath))
            {
                // If a property name and value are provided, add a custom document property.
                if (args.Length >= 3)
                {
                    string propName = args[1];
                    string propValue = args[2];

                    // Add the custom property to the workbook.
                    workbook.CustomDocumentProperties.Add(propName, propValue);

                    // Save the workbook back to the same file to persist the change.
                    workbook.Save(filePath);
                }

                // The using statement ensures Dispose() is called automatically.
                // Explicit call is optional but shown here for clarity.
                workbook.Dispose();
            }

            Console.WriteLine("Operation completed successfully.");
        }
    }
}