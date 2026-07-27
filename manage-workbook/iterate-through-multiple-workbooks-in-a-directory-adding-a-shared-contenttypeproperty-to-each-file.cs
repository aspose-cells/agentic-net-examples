using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

class AddSharedContentTypeProperty
{
    static void Main()
    {
        // Directory containing the Excel files
        string folderPath = @"C:\ExcelFiles";

        // Define the shared content type property name, value and optional type
        const string propertyName = "SharedProperty";
        const string propertyValue = "SharedValue";
        const string propertyType = "string";

        // Iterate over all .xlsx files in the directory
        foreach (string filePath in Directory.GetFiles(folderPath, "*.xlsx"))
        {
            // Load the workbook from the file
            Workbook workbook = new Workbook(filePath);

            // Add the shared content type property (if it does not already exist)
            bool exists = false;
            foreach (ContentTypeProperty prop in workbook.ContentTypeProperties)
            {
                if (prop.Name.Equals(propertyName, StringComparison.OrdinalIgnoreCase))
                {
                    exists = true;
                    break;
                }
            }

            if (!exists)
            {
                // Use the overload that specifies the property type
                workbook.ContentTypeProperties.Add(propertyName, propertyValue, propertyType);
            }

            // Save the workbook, overwriting the original file
            workbook.Save(filePath);
        }
    }
}