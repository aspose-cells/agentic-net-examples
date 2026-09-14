// Title: How to convert an Excel worksheet background image stream to a Base64 string using Aspose.Cells in C#
// AI Prompts: Write a C# method that accepts a Stream containing a worksheet background image and returns its Base64 encoded string, including null checks and stream positioning. | Demonstrate loading an Excel file with Aspose.Cells, retrieving the BackgroundImage byte array from the first worksheet, and feeding it to the conversion method to obtain a Base64 string. | Enhance the conversion routine to detect when a worksheet has no background image and throw a descriptive exception.
// Common Searches: aspocells c# get worksheet background image as base64 string | convert excel worksheet background image stream to base64 using Aspose.Cells | c# read background image bytes from Aspose.Cells worksheet and encode to base64 | how to handle missing background image when converting to base64 with Aspose.Cells
// Tags: Aspose.Cells worksheet background image to Base64 | C# stream to Base64 conversion for Excel images | extract background image bytes Aspose.Cells | encode Excel background image as Base64 in .NET | handle missing worksheet background image Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// // Loads an Excel workbook with Aspose.Cells, extracts the first worksheet's BackgroundImage byte array, converts it to a Base64 string via a helper method that processes a Stream, and prints the result while handling null checks and missing images.
public class ImageHelper
{
    /// <param name="imageStream">Stream containing the image data.</param>
    /// <returns>Base64 encoded representation of the image.</returns>
    public string ConvertBackgroundImageToBase64(Stream imageStream)
    {
        if (imageStream == null)
            throw new ArgumentNullException(nameof(imageStream));

        // Ensure the stream is positioned at the beginning.
        if (imageStream.CanSeek)
            imageStream.Position = 0;

        // Copy the input stream into a MemoryStream to obtain a byte array.
        using (var memory = new MemoryStream())
        {
            imageStream.CopyTo(memory);
            byte[] imageBytes = memory.ToArray();

            // Convert the byte array to a Base64 string.
            return Convert.ToBase64String(imageBytes);
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            // Path to the Excel file (adjust as needed).
            string excelPath = "SampleWorkbook.xlsx";

            // Verify the file exists to avoid FileNotFoundException.
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"File not found: {excelPath}");
                return;
            }

            // Load the workbook using Aspose.Cells.
            Workbook workbook = new Workbook(excelPath);

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the background image bytes, if any.
            byte[] bgBytes = sheet.BackgroundImage;

            if (bgBytes == null || bgBytes.Length == 0)
            {
                Console.WriteLine("No background image found in the worksheet.");
                return;
            }

            // Convert the background image bytes to a Base64 string.
            using (MemoryStream imgStream = new MemoryStream(bgBytes))
            {
                ImageHelper helper = new ImageHelper();
                string base64 = helper.ConvertBackgroundImageToBase64(imgStream);

                Console.WriteLine("Base64 representation of the background image:");
                Console.WriteLine(base64);
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
