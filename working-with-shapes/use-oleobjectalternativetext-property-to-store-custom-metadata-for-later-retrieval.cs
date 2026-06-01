using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class OleObjectMetadataExample
{
    static void Main()
    {
        // Path for the workbook file
        string filePath = "OleObjectMetadataDemo.xlsx";

        try
        {
            // -------------------------------------------------
            // Create a new workbook and add an OleObject
            // -------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Use a minimal valid PNG (1x1 pixel) as placeholder image data
            const string pngBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X5eUAAAAASUVORK5CYII=";
            byte[] placeholderImage = Convert.FromBase64String(pngBase64);

            // Add OleObject with the placeholder image
            int oleIndex = sheet.OleObjects.Add(5, 2, 150, 150, placeholderImage);
            OleObject ole = sheet.OleObjects[oleIndex];

            // Store custom metadata in the AlternativeText property
            ole.AlternativeText = "UserId=12345;DocumentTag=Invoice;Created=2024-10-01";

            // Save the workbook
            workbook.Save(filePath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error during workbook creation: " + ex.Message);
            return;
        }

        try
        {
            // -------------------------------------------------
            // Load the workbook and retrieve the stored metadata
            // -------------------------------------------------
            if (!File.Exists(filePath))
            {
                Console.WriteLine("Workbook file not found: " + filePath);
                return;
            }

            Workbook loadedWorkbook = new Workbook(filePath);
            OleObject loadedOle = loadedWorkbook.Worksheets[0].OleObjects[0];

            // Read the metadata from AlternativeText
            string storedMetadata = loadedOle.AlternativeText;

            // Output the retrieved metadata
            Console.WriteLine("Retrieved OleObject AlternativeText:");
            Console.WriteLine(storedMetadata);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error during workbook loading: " + ex.Message);
        }
    }
}