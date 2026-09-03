// Title: Create an MSTest unit test to add a custom document property, save as XLSX, reload, and verify persistence with Aspose.Cells for .NET
// AI Prompts: Generate an MSTest method that instantiates a Workbook, adds a custom document property, saves the file in XLSX format, reloads it, and asserts the property value remains unchanged. | Write an NUnit test case using Aspose.Cells that creates a workbook, sets a custom property, persists the workbook, loads it back, and validates the property value.
// Common Searches: how to write a unit test for custom document properties using Aspose.Cells in C# | Aspose.Cells verify custom property after saving and loading workbook | C# MSTest example for persisting custom properties in an XLSX file with Aspose.Cells | unit testing Aspose.Cells workbook property persistence .NET
// Tags: Aspose.Cells add custom document property | save workbook as XLSX Aspose.Cells | load workbook and verify property Aspose.Cells | unit test Aspose.Cells workbook properties | C# custom property persistence Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a new workbook, adds a custom document property named 'TestProperty' with the value 'HelloWorld', saves it as 'CustomPropTest.xlsx', reloads the file, checks that the property value is retained, and finally cleans up the test file.
    public class Program
    {
        private const string FilePath = "CustomPropTest.xlsx";

        public static void Main()
        {
            try
            {
                // Ensure a clean environment before execution
                if (File.Exists(FilePath))
                {
                    File.Delete(FilePath);
                }

                // Create a new workbook (creation rule)
                Workbook workbook = new Workbook();

                // Add a custom document property (custom property rule)
                workbook.CustomDocumentProperties.Add("TestProperty", "HelloWorld");

                // Save the workbook to disk (save rule)
                workbook.Save(FilePath, SaveFormat.Xlsx);

                // Verify the file was created
                if (!File.Exists(FilePath))
                {
                    Console.WriteLine($"Failed to create the file '{FilePath}'.");
                    return;
                }

                // Load the workbook from disk (load rule)
                Workbook loadedWorkbook = new Workbook(FilePath);

                // Retrieve the custom property value
                string persistedValue = loadedWorkbook.CustomDocumentProperties["TestProperty"]?.Value?.ToString();

                // Verify that the custom property persisted correctly
                if (persistedValue == "HelloWorld")
                {
                    Console.WriteLine("Custom property persisted successfully.");
                }
                else
                {
                    Console.WriteLine($"Custom property verification failed. Expected 'HelloWorld', got '{persistedValue}'.");
                }
            }
            catch (Exception ex)
            {
                // Runtime safety: log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
            finally
            {
                // Clean up the test file after execution
                try
                {
                    if (File.Exists(FilePath))
                    {
                        File.Delete(FilePath);
                    }
                }
                catch
                {
                    // Suppress any cleanup exceptions
                }
            }
        }
    }
}
