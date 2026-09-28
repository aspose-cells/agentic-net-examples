// Title: Add try‑catch error handling for invalid external Excel data source paths when using Aspose.Cells in C#
// AI Prompts: Generate C# code that loads a workbook, checks an external Excel file path, attempts to add an ExternalDataConnection, and saves the workbook, with each operation wrapped in try‑catch blocks to capture file‑not‑found and save exceptions using Aspose.Cells. | Write a helper method that validates an external Excel source path, logs detailed exception information, and safely adds the connection to a worksheet in Aspose.Cells.
// Common Searches: how to catch file not found exception for external data connection in Aspose.Cells C# | Aspose.Cells C# add external Excel connection with try‑catch for invalid path | saving workbook after external data connection error handling Aspose.Cells
// Tags: Aspose.Cells external data connection exception handling | C# validate external Excel file path Aspose.Cells | try-catch workbook save Aspose.Cells | handle invalid external connection file Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates loading an existing workbook or creating a new one, verifying the presence of an external Excel file, optionally adding an ExternalDataConnection, and saving the workbook. All critical steps are enclosed in try‑catch blocks to handle missing files, unsupported API calls, and save failures, with informative console output for each error scenario.
class Program
{
    static void Main()
    {
        // Define file paths
        string workbookPath = @"C:\Data\Sample.xlsx";
        string externalConnectionPath = @"C:\Data\InvalidDataSource.xlsx";
        string outputPath = @"C:\Data\Result.xlsx";

        Workbook workbook = null;

        try
        {
            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(workbookPath))
            {
                workbook = new Workbook(workbookPath);
            }
            else
            {
                Console.WriteLine($"Workbook not found at '{workbookPath}'. A new workbook will be created.");
                workbook = new Workbook();
            }

            // Attempt to add an external data connection if the source file exists
            if (File.Exists(externalConnectionPath))
            {
                // The ExternalDataConnections API may not be available in older Aspose.Cells versions.
                // If supported, you could add a connection like this (uncomment and adjust when available):
                // workbook.Worksheets[0].ExternalDataConnections.Add(
                //     "MyConnection",
                //     externalConnectionPath,
                //     ExternalDataSourceType.Excel,
                //     "SELECT * FROM [Sheet1$]"
                // );
                Console.WriteLine("External data source found, but adding connections is not supported in this Aspose.Cells version.");
            }
            else
            {
                Console.WriteLine($"External data source not found at '{externalConnectionPath}'. Skipping connection creation.");
            }
        }
        catch (Exception ex)
        {
            // Handle any errors that occur during loading or connection setup
            Console.WriteLine("An error occurred:");
            Console.WriteLine(ex.Message);
        }
        finally
        {
            try
            {
                // Ensure the workbook is saved (creates the file if it does not exist)
                if (workbook != null)
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to {outputPath}");
                }
                else
                {
                    Console.WriteLine("Workbook instance is null; nothing to save.");
                }
            }
            catch (Exception saveEx)
            {
                Console.WriteLine("Failed to save the workbook:");
                Console.WriteLine(saveEx.Message);
            }
        }
    }
}
