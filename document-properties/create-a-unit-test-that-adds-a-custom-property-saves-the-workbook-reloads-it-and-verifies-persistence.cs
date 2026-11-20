// Title: C# unit test for adding and persisting a custom document property with Aspose.Cells
// Common Searches: how to write a C# unit test that adds a custom document property using Aspose.Cells | Aspose.Cells verify custom document property persists after saving workbook | C# Aspose.Cells round‑trip test for custom Excel property | unit testing temporary file cleanup when saving workbook with Aspose.Cells | assert custom property value after reloading an .xlsx file in Aspose.Cells
// Tags: Aspose.Cells add custom document property | Aspose.Cells verify property persistence | Aspose.Cells save workbook as xlsx | Aspose.Cells reload workbook for validation | C# test file deletion after save

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // The example creates an in‑memory Workbook, adds a custom document property, saves the workbook to a temporary .xlsx file, reloads it, retrieves the property, and asserts that the property exists and its value matches the original, then removes the temporary file.
    public class Program
    {
        private const string PropertyName = "MyCustomProperty";
        private const string PropertyValue = "TestValue";

        public static void Main()
        {
            // Create a new workbook (in-memory)
            var workbook = new Workbook();

            // Add a custom document property
            workbook.CustomDocumentProperties.Add(PropertyName, PropertyValue);

            // Define a temporary file path for saving
            string tempFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");

            try
            {
                // Save the workbook to the temporary file
                workbook.Save(tempFilePath, SaveFormat.Xlsx);

                // Ensure the file was created before attempting to load
                if (!File.Exists(tempFilePath))
                {
                    throw new FileNotFoundException("The workbook file was not found after saving.", tempFilePath);
                }

                // Reload the workbook from the saved file
                var reloadedWorkbook = new Workbook(tempFilePath);

                // Retrieve the custom property from the reloaded workbook
                var loadedProperty = reloadedWorkbook.CustomDocumentProperties[PropertyName];

                // Verify that the property exists
                if (loadedProperty == null)
                {
                    throw new InvalidOperationException("Custom property should exist after reload.");
                }

                // Verify that the property value matches the original
                if (!PropertyValue.Equals(loadedProperty.Value?.ToString()))
                {
                    throw new InvalidOperationException($"Custom property value mismatch. Expected: {PropertyValue}, Actual: {loadedProperty.Value}");
                }

                Console.WriteLine("Custom property persisted successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
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
                    catch
                    {
                        // Ignored - cleanup failure should not crash the program
                    }
                }
            }
        }
    }
}
