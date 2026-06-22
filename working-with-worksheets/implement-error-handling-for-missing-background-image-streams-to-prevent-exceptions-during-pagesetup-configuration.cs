using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class PageSetupBackgroundImageHandling
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            PageSetup pageSetup = worksheet.PageSetup;

            // Attempt to load an image file into a byte array
            byte[] imageData = null;
            string imagePath = "header.png"; // Path to the background image

            try
            {
                // Verify that the file exists before opening
                if (File.Exists(imagePath))
                {
                    using (FileStream fileStream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                        fileStream.CopyTo(memoryStream);
                        imageData = memoryStream.ToArray();
                    }
                }
                else
                {
                    Console.WriteLine($"Image file \"{imagePath}\" not found. Header picture will be skipped.");
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected I/O errors and continue without setting the picture
                Console.WriteLine($"Error loading image: {ex.Message}. Header picture will be skipped.");
            }

            // Only set the picture if valid image data is available
            if (imageData != null && imageData.Length > 0)
            {
                try
                {
                    // Set the picture in the center section of the header (section = 1)
                    // Parameters: isFirstPage = false, isEvenPage = false, isHeader = true, section = 1, imageData
                    pageSetup.SetPicture(false, false, true, 1, imageData);

                    // Use the special code "&G" to display the picture in the header
                    pageSetup.SetHeader(1, "&G");

                    Console.WriteLine("Header picture set successfully.");
                }
                catch (Exception ex)
                {
                    // Handle any exceptions thrown by SetPicture (e.g., unsupported format)
                    Console.WriteLine($"Error setting header picture: {ex.Message}");
                }
            }

            // Save the workbook to verify that no exception occurs during save
            string outputPath = "PageSetupBackgroundImageDemo.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
    }
}