using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsUtilities
{
    /// <summary>
    /// Provides helper methods for working with worksheet background images.
    /// </summary>
    public static class BackgroundImageHelper
    {
        /// <summary>
        /// Reads an image stream (e.g., a worksheet background image) and returns its Base64 representation.
        /// </summary>
        /// <param name="imageStream">The input stream containing the image data.</param>
        /// <returns>A Base64 encoded string of the image.</returns>
        public static string ConvertStreamToBase64(Stream imageStream)
        {
            if (imageStream == null)
                throw new ArgumentNullException(nameof(imageStream));

            // Ensure the stream is positioned at the beginning.
            if (imageStream.CanSeek)
                imageStream.Position = 0;

            // Copy the stream into a memory buffer and convert to Base64.
            using (MemoryStream memory = new MemoryStream())
            {
                imageStream.CopyTo(memory);
                return Convert.ToBase64String(memory.ToArray());
            }
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                const string imagePath = "background.jpg";

                // Prevent FileNotFoundException.
                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"File not found: {imagePath}");
                    return;
                }

                // Create a new workbook and set the background image.
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                worksheet.BackgroundImage = File.ReadAllBytes(imagePath);

                // Convert the background image bytes to a Base64 string.
                using (MemoryStream stream = new MemoryStream(worksheet.BackgroundImage))
                {
                    string base64 = BackgroundImageHelper.ConvertStreamToBase64(stream);
                    Console.WriteLine("Base64 string:");
                    Console.WriteLine(base64);
                }
            }
            catch (Exception ex)
            {
                // Runtime safety.
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}