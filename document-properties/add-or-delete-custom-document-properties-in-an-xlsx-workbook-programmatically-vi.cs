using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsCustomPropertiesDemo
{
    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook(); // create

            // ---------- Add custom document properties ----------
            // String property
            workbook.CustomDocumentProperties.Add("ProjectName", "Alpha");
            // Integer property
            workbook.CustomDocumentProperties.Add("Revision", 3);
            // DateTime property
            workbook.CustomDocumentProperties.Add("LastReviewed", DateTime.Now);
            // Boolean property
            workbook.CustomDocumentProperties.Add("Approved", true);
            // Double (float) property
            workbook.CustomDocumentProperties.Add("Score", 4.75);

            // Save the workbook with the added properties
            workbook.Save("CustomPropertiesAdded.xlsx"); // save

            // ---------- Load the existing workbook ----------
            Workbook loadedWorkbook = new Workbook("CustomPropertiesAdded.xlsx"); // load

            // ---------- Delete a custom document property ----------
            // Remove the property named "Approved"
            loadedWorkbook.CustomDocumentProperties.Remove("Approved");

            // Optionally, verify removal
            Console.WriteLine("Properties after removal:");
            foreach (DocumentProperty prop in loadedWorkbook.CustomDocumentProperties)
            {
                Console.WriteLine($"{prop.Name}: {prop.Value}");
            }

            // Save the workbook after deletion
            loadedWorkbook.Save("CustomPropertiesModified.xlsx"); // save
        }
    }
}