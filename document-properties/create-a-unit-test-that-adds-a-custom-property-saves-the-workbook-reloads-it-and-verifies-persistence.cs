// Title: C# unit test with Aspose.Cells that adds a custom document property, saves to XLSX, reloads, and verifies persistence
// Common Searches: how to write a unit test for custom document properties using Aspose.Cells in C# | Aspose.Cells C# verify custom property remains after workbook save and load | unit testing workbook custom properties with temporary file in Aspose.Cells | C# Aspose.Cells test to ensure custom document property persists in saved Excel file | example of adding and retrieving custom document property in Aspose.Cells unit test
// Tags: Aspose.Cells custom document property testing | C# workbook save reload verification | Excel custom property persistence Aspose.Cells | temporary file cleanup Aspose.Cells unit test | validate custom property value after load

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // Demonstrates a C# unit test that creates a workbook, adds a custom document property, saves the file to a temporary XLSX, reloads the workbook, retrieves the property, asserts the value matches the original, and cleans up the temporary file.
    public class Program
    {
        private const string PropertyName = "MyCustomProperty";
        private const string PropertyValue = "TestValue";

        public static void Main()
        {
            // Create a new workbook in memory
            Workbook workbook = new Workbook();

            // Add a custom document property
            workbook.CustomDocumentProperties.Add(PropertyName, PropertyValue);

            // Define a temporary file path for saving
            string tempFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");

            try
            {
                // Save the workbook to the temporary file
                workbook.Save(tempFilePath, SaveFormat.Xlsx);

                // Ensure the file exists before attempting to load
                if (!File.Exists(tempFilePath))
                {
                    Console.WriteLine("Error: Temporary file was not created.");
                    return;
                }

                // Load the workbook from the saved file
                Workbook loadedWorkbook = new Workbook(tempFilePath);

                // Retrieve the custom property from the loaded workbook
                string loadedValue = loadedWorkbook.CustomDocumentProperties[PropertyName]?.Value?.ToString();

                // Verify that the property value persisted correctly
                if (PropertyValue == loadedValue)
                {
                    Console.WriteLine("Success: The custom property value was persisted after save and load.");
                }
                else
                {
                    Console.WriteLine($"Failure: Expected '{PropertyValue}' but got '{loadedValue ?? "null"}'.");
                }
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"Exception occurred: {ex.Message}");
            }
            finally
            {
                // Clean up the temporary file
                if (File.Exists(tempFilePath))
                {
                    try
                    {
                        File.Delete(tempFilePath);
                    }
                    catch (Exception deleteEx)
                    {
                        Console.WriteLine($"Failed to delete temporary file: {deleteEx.Message}");
                    }
                }
            }
        }
    }
}
