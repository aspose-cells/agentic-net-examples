// Title: Select and load an external Excel workbook via a ribbon button using Aspose.Cells in C#
// Common Searches: how to open a file picker from a ribbon button and load an Excel workbook with Aspose.Cells in C# | C# Aspose.Cells example for loading a user‑selected external workbook in a ribbon UI | validate and open an external Excel file chosen through a ribbon command using Aspose.Cells | retrieve the first worksheet name after a user picks an Excel file in a .NET ribbon button | exception handling when loading an external workbook selected via a ribbon button with Aspose.Cells
// Tags: Aspose.Cells load workbook from user selected file | C# ribbon button file selection Aspose.Cells | validate external Excel file existence Aspose.Cells | access first worksheet name Aspose.Cells | exception handling workbook loading Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace MyApp
{
    // The example shows how to attach a handler to a ribbon button that prompts the user for an Excel file path, verifies the file exists, loads the workbook with Aspose.Cells, accesses the first worksheet, and outputs its name, all with proper exception handling.
    public class RibbonButtonHandler
    {
        // This method should be attached to the ribbon button click event.
        public void OnSelectExternalDataSource(object sender, EventArgs e)
        {
            try
            {
                // Prompt the user to enter the full path of the external Excel file.
                Console.WriteLine("Enter the full path of the external Excel file:");
                string externalFilePath = Console.ReadLine();

                if (string.IsNullOrWhiteSpace(externalFilePath))
                {
                    Console.WriteLine("No file path was entered. Operation cancelled.");
                    return;
                }

                // Verify that the file exists before attempting to load it.
                if (!File.Exists(externalFilePath))
                {
                    Console.WriteLine($"File not found: {externalFilePath}");
                    return;
                }

                // Load the selected external data source workbook using Aspose.Cells.
                Workbook externalWorkbook = new Workbook(externalFilePath);

                // Example: Access the first worksheet of the external workbook.
                Worksheet externalSheet = externalWorkbook.Worksheets[0];

                // For demonstration, show the name of the first worksheet.
                Console.WriteLine($"Loaded external worksheet: {externalSheet.Name}");
            }
            catch (Exception ex)
            {
                // Log or display the exception details for troubleshooting.
                Console.WriteLine($"An error occurred while loading the external data source: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                RibbonButtonHandler handler = new RibbonButtonHandler();
                // Simulate ribbon button click; sender and EventArgs can be null.
                handler.OnSelectExternalDataSource(null, EventArgs.Empty);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
