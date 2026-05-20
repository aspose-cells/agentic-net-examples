using System;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsTests
{
    public class CustomPropertyPersistenceTests
    {
        public void Run()
        {
            // Arrange: create a workbook and add a custom property
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            const string propertyName = "TestProp";
            const string propertyValue = "TestValue";
            sheet.CustomProperties.Add(propertyName, propertyValue);

            // Temporary file for saving the workbook
            string tempFile = Path.Combine(Path.GetTempPath(), $"CustomProp_{Guid.NewGuid()}.xlsx");

            try
            {
                // Act: save the workbook
                workbook.Save(tempFile);

                // Ensure the file exists before loading
                if (!File.Exists(tempFile))
                    throw new FileNotFoundException("Saved workbook not found.", tempFile);

                // Load the workbook back
                Workbook loadedWorkbook = new Workbook(tempFile);
                Worksheet loadedSheet = loadedWorkbook.Worksheets[0];

                // Retrieve the custom property
                CustomProperty loadedProperty = loadedSheet.CustomProperties
                    .FirstOrDefault(p => p.Name == propertyName);

                // Assert: property exists and value matches
                if (loadedProperty == null)
                    throw new InvalidOperationException("Custom property was not found after loading the workbook.");

                if (!propertyValue.Equals(loadedProperty.Value?.ToString()))
                    throw new InvalidOperationException("Custom property value did not persist correctly.");

                Console.WriteLine("Custom property persisted successfully.");
            }
            catch (Exception ex)
            {
                // Runtime safety: log and rethrow
                Console.WriteLine($"Error: {ex.Message}");
                throw;
            }
            finally
            {
                // Cleanup: delete temporary file if it exists
                if (File.Exists(tempFile))
                {
                    try { File.Delete(tempFile); } catch { /* ignore cleanup errors */ }
                }
            }
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                var test = new CustomPropertyPersistenceTests();
                test.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test failed: {ex.Message}");
            }
        }
    }
}