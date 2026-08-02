using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsDemo
{
    public class Program
    {
        private const string TestFile = "ContentTypePropertyIsNillableTest.xlsx";

        public static void Main()
        {
            try
            {
                // Create a new workbook and add a content type property
                var workbook = new Workbook();
                workbook.ContentTypeProperties.Add("Admin", "Aspose", "text");

                // Retrieve the property and set IsNillable to true
                ContentTypeProperty property = workbook.ContentTypeProperties["Admin"];
                property.IsNillable = true;

                // Save the workbook to a file
                workbook.Save(TestFile);

                // Ensure the file exists before loading
                if (!File.Exists(TestFile))
                    throw new FileNotFoundException($"The file '{TestFile}' was not found after saving.");

                // Load the workbook from the saved file
                var loadedWorkbook = new Workbook(TestFile);
                ContentTypeProperty loadedProperty = loadedWorkbook.ContentTypeProperties["Admin"];

                // Verify that IsNillable remains true after loading
                if (loadedProperty.IsNillable)
                {
                    Console.WriteLine("IsNillable flag is true after loading. Test passed.");
                }
                else
                {
                    Console.WriteLine("IsNillable flag is false after loading. Test failed.");
                }
            }
            catch (Exception ex)
            {
                // Runtime safety: report any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}