using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Verify that the Word document exists
            const string wordPath = "Sample.docx";
            if (!File.Exists(wordPath))
                throw new FileNotFoundException($"Word file not found: {wordPath}");

            // Load the Word document to be embedded
            byte[] wordData = File.ReadAllBytes(wordPath);

            // Verify that the custom icon image exists
            const string iconPath = "CustomIcon.png";
            if (!File.Exists(iconPath))
                throw new FileNotFoundException($"Icon image not found: {iconPath}");

            // Load the custom icon image that will represent the OLE object
            byte[] iconImage = File.ReadAllBytes(iconPath);

            // Add an OLE object placeholder using the icon image as the preview picture
            // Parameters: topRow, leftColumn, height (px), width (px), imageData
            int oleIndex = worksheet.OleObjects.Add(5, 2, 100, 100, iconImage);
            OleObject oleObject = worksheet.OleObjects[oleIndex];

            // Embed the Word document into the OLE object
            // linkToFile = false (embed), displayAsIcon = true, updateIcon = false
            oleObject.SetEmbeddedObject(
                linkToFile: false,
                objectData: wordData,
                sourceFileName: wordPath,
                displayAsIcon: true,
                label: "Word Document",
                updateIcon: false);

            // Ensure the object is displayed as an icon
            oleObject.DisplayAsIcon = true;

            // Set the custom icon image file path (the image will be used as the display icon)
            oleObject.ImageSourceFullName = iconPath;

            // Save the workbook with the embedded Word OLE object
            workbook.Save("WordOleWithCustomIcon.xlsx");
        }
        catch (FileNotFoundException fnfEx)
        {
            Console.WriteLine($"File error: {fnfEx.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}