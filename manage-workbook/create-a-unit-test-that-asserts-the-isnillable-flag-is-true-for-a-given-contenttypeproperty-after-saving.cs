using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTests
{
    class Program
    {
        private const string PropertyName = "Admin";
        private const string PropertyValue = "Aspose";
        private const string PropertyType = "text";

        static void Main()
        {
            string tempFile = null;

            try
            {
                // Create a new workbook and add a content type property
                var workbook = new Workbook();
                workbook.ContentTypeProperties.Add(PropertyName, PropertyValue, PropertyType);

                // Retrieve the added property and set IsNillable flag
                var property = workbook.ContentTypeProperties[PropertyName];
                property.IsNillable = true;

                // Save the workbook to a temporary file
                tempFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");
                workbook.Save(tempFile);

                // Ensure the file was created before attempting to load it
                if (!File.Exists(tempFile))
                {
                    Console.WriteLine("Temporary file was not created.");
                    return;
                }

                // Load the saved workbook
                var loadedWorkbook = new Workbook(tempFile);
                var loadedProperty = loadedWorkbook.ContentTypeProperties[PropertyName];

                // Verify that IsNillable remains true after loading
                if (loadedProperty.IsNillable)
                {
                    Console.WriteLine("Test passed: IsNillable is true after saving and loading.");
                }
                else
                {
                    Console.WriteLine("Test failed: IsNillable is false after saving and loading.");
                }
            }
            catch (Exception ex)
            {
                // Runtime safety: capture any unexpected errors
                Console.WriteLine($"Exception: {ex.Message}");
            }
            finally
            {
                // Cleanup: delete the temporary file if it exists
                if (!string.IsNullOrEmpty(tempFile) && File.Exists(tempFile))
                {
                    try
                    {
                        File.Delete(tempFile);
                    }
                    catch
                    {
                        // Suppress any cleanup errors
                    }
                }
            }
        }
    }
}