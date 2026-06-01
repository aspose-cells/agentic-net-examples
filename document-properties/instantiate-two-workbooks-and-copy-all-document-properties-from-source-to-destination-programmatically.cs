using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsDocumentPropertiesCopy
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create source workbook and set some built‑in and custom properties
                Workbook sourceWorkbook = new Workbook();
                sourceWorkbook.BuiltInDocumentProperties["Author"].Value = "John Smith";
                sourceWorkbook.BuiltInDocumentProperties["Title"].Value = "Sample Source Workbook";
                sourceWorkbook.CustomDocumentProperties.Add("ReviewedBy", "Jane Doe");
                sourceWorkbook.CustomDocumentProperties.Add("Revision", 3);

                // Create destination workbook (empty)
                Workbook destinationWorkbook = new Workbook();

                // ----- Copy Built‑in Document Properties -----
                foreach (DocumentProperty srcProp in sourceWorkbook.BuiltInDocumentProperties)
                {
                    // Ensure the destination has the same property and assign its value
                    destinationWorkbook.BuiltInDocumentProperties[srcProp.Name].Value = srcProp.Value;
                }

                // ----- Copy Custom Document Properties -----
                foreach (DocumentProperty srcProp in sourceWorkbook.CustomDocumentProperties)
                {
                    // If the property already exists, update its value; otherwise add it
                    if (destinationWorkbook.CustomDocumentProperties.Contains(srcProp.Name))
                    {
                        destinationWorkbook.CustomDocumentProperties[srcProp.Name].Value = srcProp.Value;
                    }
                    else
                    {
                        // Convert the value to string to match the overload that accepts (string, string)
                        destinationWorkbook.CustomDocumentProperties.Add(srcProp.Name, srcProp.Value?.ToString() ?? string.Empty);
                    }
                }

                // Save both workbooks to verify the copy
                sourceWorkbook.Save("SourceWorkbook.xlsx");
                destinationWorkbook.Save("DestinationWorkbook.xlsx");

                Console.WriteLine("Document properties copied successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}