using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsLoadWarningsDemo
{
    // Custom warning callback that stores warnings for later use
    public class CustomWarningCallback : IWarningCallback
    {
        // List to keep all received warnings
        public List<WarningInfo> Warnings { get; } = new List<WarningInfo>();

        // This method is called by Aspose.Cells when a warning occurs
        public void Warning(WarningInfo warningInfo)
        {
            // Store the warning
            Warnings.Add(warningInfo);
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the Excel file to be loaded
                string filePath = "input.xlsx";

                // Ensure the input file exists to avoid FileNotFoundException
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
                    return;
                }

                // Create LoadOptions and assign the custom warning callback
                LoadOptions loadOptions = new LoadOptions
                {
                    WarningCallback = new CustomWarningCallback()
                    // LightCellsDataHandler is optional; omitted to avoid API mismatches
                };

                // Load the workbook using the specified options
                Workbook workbook = new Workbook(filePath, loadOptions);

                // Output all collected warnings after loading
                Console.WriteLine("Load Warnings:");
                foreach (WarningInfo warning in ((CustomWarningCallback)loadOptions.WarningCallback).Warnings)
                {
                    Console.WriteLine($"- Type: {warning.WarningType}, Description: {warning.Description}");
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}