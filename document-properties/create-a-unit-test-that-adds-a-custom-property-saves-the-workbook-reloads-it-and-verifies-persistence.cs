using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsTests
{
    public class CustomPropertyPersistenceTests
    {
        private const string PropertyName = "TestProperty";
        private const string PropertyValue = "HelloWorld";

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
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a custom property to the worksheet
            CustomPropertyCollection customProps = sheet.CustomProperties;
            customProps.Add(PropertyName, PropertyValue);

            // Save the workbook to a temporary file
            string tempFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");
            workbook.Save(tempFile, SaveFormat.Xlsx);

            // Ensure the file exists before loading
            if (!File.Exists(tempFile))
                throw new FileNotFoundException("Saved workbook file not found.", tempFile);

            // Load the workbook from the saved file
            Workbook loadedWorkbook = new Workbook(tempFile);
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
            CustomPropertyCollection loadedCustomProps = loadedSheet.CustomProperties;

            // Retrieve the property and verify its value
            bool found = false;
            foreach (CustomProperty prop in loadedCustomProps)
            {
                if (prop.Name == PropertyName)
                {
                    found = true;
                    if (prop.Value?.ToString() != PropertyValue)
                        throw new InvalidOperationException("Custom property value mismatch after reload.");
                    break;
                }
            }

            if (!found)
                throw new InvalidOperationException($"Custom property '{PropertyName}' was not found after reload.");

            // Clean up the temporary file
            try
            {
                if (File.Exists(tempFile))
                    File.Delete(tempFile);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Unable to delete temporary file. {ex.Message}");
            }
        }
    }
}