// Title: C# Unit Test for Verifying Custom Worksheet Property Persistence with Aspose.Cells
// Description: Demonstrates how to add a custom property to a worksheet, save the workbook to a temporary .xlsx file, reload it, and assert that the property and its value are retained, with proper cleanup.
// Keywords: Aspose.Cells | C# | .NET | custom worksheet property | unit test | MSTest | NUnit | xUnit | property persistence | Excel metadata | temporary file handling | load and save workbook | automated testing
// Common Searches: Aspose.Cells unit test custom property C# | verify worksheet custom property after save | C# test Excel custom metadata persistence | how to assert custom property in Aspose.Cells workbook | temporary file cleanup in Aspose.Cells tests
// Developer Intent: Write a C# unit test that adds a custom worksheet property, saves the workbook, reloads it, and confirms the property persists.
// Use Cases: Automated regression testing to ensure custom metadata is not lost during file I/O. | Continuous‑integration validation of property round‑trip across Excel formats. | Sample code for developers needing to verify custom properties in CI pipelines.
// AI Prompts: Generate an MSTest method in C# using Aspose.Cells that adds a custom worksheet property, saves to a temp .xlsx, reloads, asserts the value, and deletes the file. | Create an NUnit test for Aspose.Cells that verifies a custom property survives a save‑load cycle, including exception handling and cleanup. | Provide an xUnit test example that checks custom worksheet property persistence with Aspose.Cells, using fluent assertions and temporary file management.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

// Demonstrates how to add a custom property to a worksheet, save the workbook to a temporary .xlsx file, reload it, and assert that the property and its value are retained, with proper cleanup.
class Program
{
    static void Main()
    {
        string tempFile = null;

        try
        {
            // Create a new workbook and add a custom property to the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            CustomPropertyCollection customProps = sheet.CustomProperties;
            const string propName = "TestProp";
            const string propValue = "HelloWorld";
            customProps.Add(propName, propValue);

            // Save the workbook to a temporary file
            tempFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");
            workbook.Save(tempFile);

            // Ensure the file was created before attempting to load it
            if (!File.Exists(tempFile))
                throw new FileNotFoundException("The workbook file was not found after saving.", tempFile);

            // Reload the workbook from the saved file
            Workbook loadedWorkbook = new Workbook(tempFile);
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
            CustomProperty loadedProp = loadedSheet.CustomProperties[propName];

            // Verify that the custom property exists and its value is preserved
            if (loadedProp == null)
                throw new Exception("Custom property should exist after reload.");

            if (!propValue.Equals(loadedProp.Value?.ToString()))
                throw new Exception($"Custom property value mismatch. Expected: {propValue}, Actual: {loadedProp.Value}");

            Console.WriteLine("Custom property persisted successfully.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
        finally
        {
            // Clean up the temporary file
            if (!string.IsNullOrEmpty(tempFile) && File.Exists(tempFile))
            {
                try
                {
                    File.Delete(tempFile);
                }
                catch
                {
                    // Suppress any exceptions during cleanup
                }
            }
        }
    }
}
