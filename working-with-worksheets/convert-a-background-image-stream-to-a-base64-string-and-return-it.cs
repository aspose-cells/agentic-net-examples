// Title: C# – Convert Worksheet Background Image Stream to Base64 String
// Description: A reusable C# helper that validates a non‑null image Stream, rewinds it if possible, reads all bytes via a MemoryStream, and returns the Base64 representation. Includes a demo that loads a JPEG file, encodes it, and prints the result—ideal for embedding background pictures in Excel files with Aspose.Cells.
// Keywords: C# base64 image stream | worksheet background image | Aspose.Cells background picture | convert stream to base64 | image to base64 C# | Excel background image encoding | file stream to base64
// Common Searches: convert image stream to base64 c# | aspocells background image base64 | c# read file stream and get base64 | embed base64 background image in Excel using Aspose.Cells | c# encode worksheet background picture
// Developer Intent: Generate a Base64‑encoded string from a worksheet background image stream for use with Aspose.Cells.
// Use Cases: Encode a background picture to Base64 before assigning it to a worksheet via Aspose.Cells. | Transmit the Base64 string of a worksheet background image through a REST API for client‑side rendering. | Persist the Base64 representation of a worksheet background image in a database for later reuse.
// AI Prompts: Write a C# method that accepts any image Stream and returns its Base64 string, including null checks and stream positioning. | Show how to read a JPEG file, convert it to Base64 with the helper, and set it as a worksheet background using Aspose.Cells. | Refactor ConvertBackgroundImageStreamToBase64 into an async method using async/await.

using System;
using System.IO;

// A reusable C# helper that validates a non‑null image Stream, rewinds it if possible, reads all bytes via a MemoryStream, and returns the Base64 representation. Includes a demo that loads a JPEG file, encodes it, and prints the result—ideal for embedding background pictures in Excel files with Aspose.Cells.
public static class ImageHelper
{
    /// <param name="imageStream">The input stream containing the image data.</param>
    /// <returns>Base64 encoded string of the image.</returns>
    public static string ConvertBackgroundImageStreamToBase64(Stream imageStream)
    {
        if (imageStream == null)
            throw new ArgumentNullException(nameof(imageStream));

        // Ensure the stream is positioned at the beginning
        if (imageStream.CanSeek)
            imageStream.Seek(0, SeekOrigin.Begin);

        // Read all bytes from the stream
        using (MemoryStream ms = new MemoryStream())
        {
            imageStream.CopyTo(ms);
            byte[] imageBytes = ms.ToArray();

            // Convert the byte array to a Base64 string
            return Convert.ToBase64String(imageBytes);
        }
    }

    // Example usage
    public static void Demo()
    {
        // Assume we have an image file that will be used as a worksheet background
        string imagePath = "background.jpg";

        if (!File.Exists(imagePath))
        {
            Console.Error.WriteLine($"File not found: {imagePath}");
            return;
        }

        try
        {
            using (FileStream fs = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
            {
                string base64 = ConvertBackgroundImageStreamToBase64(fs);
                Console.WriteLine("Base64 string:");
                Console.WriteLine(base64);
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error processing image: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            ImageHelper.Demo();
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}
