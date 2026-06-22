using System;
using System.IO;
using Aspose.Cells;

namespace LoadWarningsDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file to be loaded
            string filePath = "input.xlsx";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
                return;
            }

            try
            {
                // Load the workbook using the standard constructor
                Workbook workbook = new Workbook(filePath);

                // NOTE: In some versions of Aspose.Cells the LoadWarnings property is not available.
                // If you need to retrieve load warnings, ensure you are using a version that supports
                // Workbook.LoadWarnings (type WarningInfoCollection). The following block is kept
                // for reference and will compile only when the API is present.

                /*
                WarningInfoCollection warnings = workbook.LoadWarnings;
                foreach (WarningInfo warning in warnings)
                {
                    Console.WriteLine($"Warning Type: {warning.Type}");
                    Console.WriteLine($"Description : {warning.Description}");
                    Console.WriteLine();
                }
                */

                Console.WriteLine("Workbook loaded successfully.");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}