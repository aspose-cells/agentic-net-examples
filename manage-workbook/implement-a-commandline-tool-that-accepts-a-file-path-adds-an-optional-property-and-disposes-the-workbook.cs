using System;
using Aspose.Cells;

namespace AsposeCellsCommandLineTool
{
    class Program
    {
        // Entry point of the command‑line application
        static void Main(string[] args)
        {
            // Expect at least one argument: the path to the Excel file
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: AsposeCellsCommandLineTool <filePath> [propertyName propertyValue]");
                return;
            }

            string filePath = args[0];

            // Validate that the file exists before attempting to load it
            if (!System.IO.File.Exists(filePath))
            {
                Console.WriteLine($"Error: File not found - {filePath}");
                return;
            }

            // Load the workbook using the string constructor (load rule)
            Workbook workbook = new Workbook(filePath);

            // If a property name and value are supplied, add a custom document property
            if (args.Length >= 3)
            {
                string propName = args[1];
                string propValue = args[2];

                // Add the custom property (no specific rule exists for this operation)
                workbook.CustomDocumentProperties.Add(propName, propValue);

                // Save the workbook back to the same file (save rule)
                workbook.Save(filePath);
                Console.WriteLine($"Added custom property '{propName}' with value '{propValue}' and saved workbook.");
            }
            else
            {
                Console.WriteLine("No custom property specified. Workbook loaded and will be disposed.");
            }

            // Dispose the workbook to release unmanaged resources (dispose rule)
            workbook.Dispose();
        }
    }
}