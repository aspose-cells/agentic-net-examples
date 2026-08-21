// Title: Aspose.Cells .NET Unit Test: Add Custom Worksheet Property, Save, Reload, Verify Persistence
// Description: C# example that builds a Workbook, inserts a custom property into the first worksheet, saves the workbook to a temporary .xlsx file, reloads it, and asserts that the property name and value are unchanged. The test also removes the temporary file after verification.
// Keywords: Aspose.Cells | C# unit test | custom worksheet property | Excel custom property persistence | save and reload workbook | temporary file handling | MSTest example | xUnit test | NUnit scenario | .NET Excel automation
// Common Searches: how to unit test custom worksheet properties with Aspose.Cells | verify Excel custom property persists after save .NET | Aspose.Cells C# example add and read custom properties | unit testing Excel metadata using temporary files | Aspose.Cells unit test for document property persistence
// Developer Intent: Write an automated test that adds a custom property to a worksheet, saves the workbook, reloads it, and confirms the property remains intact.
// Use Cases: Continuous‑integration validation that generated reports keep custom metadata. | Regression testing for services that embed business identifiers in Excel worksheets. | Quality‑gate checks ensuring document properties survive serialization across environments.
// AI Prompts: Generate an MSTest method that adds a custom worksheet property with Aspose.Cells, saves to a temp file, reloads the workbook, and asserts the property value. | Provide an xUnit test snippet that creates a Workbook, inserts a custom property, persists the file, reads it back, and uses Assert.Equal to verify the value. | Write a NUnit test that adds a custom property to the first sheet, saves as .xlsx, loads the file, and confirms the property exists with the expected string.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsTests
{
    // C# example that builds a Workbook, inserts a custom property into the first worksheet, saves the workbook to a temporary .xlsx file, reloads it, and asserts that the property name and value are unchanged. The test also removes the temporary file after verification.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add a custom property to the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            CustomPropertyCollection customProps = sheet.CustomProperties;
            const string propertyName = "TestProperty";
            const string propertyValue = "TestValue";

            try
            {
                customProps.Add(propertyName, propertyValue);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error adding custom property: {ex.Message}");
                return;
            }

            // Define a temporary file path for saving the workbook
            string tempFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");

            try
            {
                // Save the workbook
                workbook.Save(tempFilePath, SaveFormat.Xlsx);

                // Verify the file was created before attempting to load
                if (!File.Exists(tempFilePath))
                {
                    Console.WriteLine("Failed to create the temporary workbook file.");
                    return;
                }

                // Load the workbook from the saved file
                Workbook loadedWorkbook = new Workbook(tempFilePath);
                Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
                CustomPropertyCollection loadedCustomProps = loadedSheet.CustomProperties;

                // Retrieve the property by name
                CustomProperty loadedProperty = null;
                foreach (CustomProperty prop in loadedCustomProps)
                {
                    if (prop.Name == propertyName)
                    {
                        loadedProperty = prop;
                        break;
                    }
                }

                // Validate that the property exists and its value matches the original
                if (loadedProperty == null)
                {
                    Console.WriteLine($"Custom property '{propertyName}' was not found after loading.");
                }
                else if (!propertyValue.Equals(loadedProperty.Value?.ToString()))
                {
                    Console.WriteLine($"Custom property value mismatch. Expected: '{propertyValue}', Actual: '{loadedProperty.Value}'.");
                }
                else
                {
                    Console.WriteLine("Custom property persisted correctly after save and load.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during workbook processing: {ex.Message}");
            }
            finally
            {
                // Cleanup: delete the temporary file if it exists
                try
                {
                    if (File.Exists(tempFilePath))
                    {
                        File.Delete(tempFilePath);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to delete temporary file: {ex.Message}");
                }
            }
        }
    }
}
