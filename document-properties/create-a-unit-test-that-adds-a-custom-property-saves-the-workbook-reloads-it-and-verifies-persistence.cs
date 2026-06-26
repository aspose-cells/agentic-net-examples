using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsTests
{
    public class CustomPropertyPersistenceTests
    {
        private const string PropertyName = "TestProp";
        private const string PropertyValue = "TestValue";

        public static void Main()
        {
            try
            {
                new CustomPropertyPersistenceTests().RunTest();
                Console.WriteLine("Test passed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test failed: {ex.Message}");
            }
        }

        public void RunTest()
        {
            string tempFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");

            try
            {
                // Create a new workbook and add a custom property to the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.CustomProperties.Add(PropertyName, PropertyValue);

                // Save the workbook
                workbook.Save(tempFile);

                // Verify the file exists before loading
                if (!File.Exists(tempFile))
                    throw new FileNotFoundException("Saved workbook not found.", tempFile);

                // Load the workbook from the saved file
                Workbook loadedWorkbook = new Workbook(tempFile);
                Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
                CustomPropertyCollection loadedProps = loadedSheet.CustomProperties;

                // Find the custom property by name
                CustomProperty foundProp = null;
                foreach (CustomProperty prop in loadedProps)
                {
                    if (prop.Name == PropertyName)
                    {
                        foundProp = prop;
                        break;
                    }
                }

                // Validate that the property exists and its value matches
                if (foundProp == null)
                    throw new InvalidOperationException($"Custom property '{PropertyName}' was not found after loading.");

                if (!PropertyValue.Equals(foundProp.Value?.ToString()))
                    throw new InvalidOperationException("Custom property value does not match after loading.");
            }
            finally
            {
                // Cleanup: delete the temporary file if it exists
                if (File.Exists(tempFile))
                {
                    try
                    {
                        File.Delete(tempFile);
                    }
                    catch
                    {
                        // Ignore cleanup errors
                    }
                }
            }
        }
    }
}