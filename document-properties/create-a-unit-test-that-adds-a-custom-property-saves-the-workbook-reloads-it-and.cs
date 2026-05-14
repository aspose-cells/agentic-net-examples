using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTests
{
    class Program
    {
        private const string PropertyName = "TestProp";
        private const string PropertyValue = "TestValue";

        static void Main()
        {
            try
            {
                // ---------- Create ----------
                var workbook = new Workbook();
                var worksheet = workbook.Worksheets[0];

                // Add a custom property to the worksheet
                worksheet.CustomProperties.Add(PropertyName, PropertyValue);

                // ---------- Save ----------
                string tempFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");
                workbook.Save(tempFile);
                workbook.Dispose();

                // ---------- Load ----------
                var loadedWorkbook = new Workbook(tempFile);
                var loadedWorksheet = loadedWorkbook.Worksheets[0];

                // ---------- Verify ----------
                var foundProperty = loadedWorksheet.CustomProperties[PropertyName];

                if (foundProperty == null)
                {
                    throw new Exception($"Custom property '{PropertyName}' was not found after loading.");
                }

                if (foundProperty.Value?.ToString() != PropertyValue)
                {
                    throw new Exception("Custom property value does not match after loading.");
                }

                Console.WriteLine("Custom property persisted successfully.");

                // Cleanup
                loadedWorkbook.Dispose();
                File.Delete(tempFile);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test failed: {ex.Message}");
            }
        }
    }
}