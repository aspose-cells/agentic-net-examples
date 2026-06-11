using System;
using System.IO;
using Aspose.Cells;

class SetWorksheetBackground
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Path to the PNG file that will be used as background
            string pngFilePath = "background.png";

            // Verify that the background image file exists
            if (File.Exists(pngFilePath))
            {
                // Load the PNG file into a byte array
                byte[] backgroundBytes;
                using (FileStream fileStream = new FileStream(pngFilePath, FileMode.Open, FileAccess.Read))
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    fileStream.CopyTo(memoryStream);
                    backgroundBytes = memoryStream.ToArray();
                }

                // Assign the byte array to the worksheet's BackgroundImage property
                worksheet.BackgroundImage = backgroundBytes;
            }
            else
            {
                Console.WriteLine($"Background image file not found: {pngFilePath}");
            }

            // Save the workbook (with or without background image)
            string outputPath = "WorkbookWithBackground.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}