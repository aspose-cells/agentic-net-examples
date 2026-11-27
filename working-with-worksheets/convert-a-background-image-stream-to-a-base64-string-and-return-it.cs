// Title: Convert an Excel worksheet background image to a Base64 string using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# read worksheet background image and get Base64 | how to extract background picture from Excel file stream with Aspose.Cells | convert Excel sheet background image to Base64 in .NET Core | retrieve background image bytes from workbook using Aspose.Cells API | C# get background image from first worksheet and encode to Base64
// Tags: Aspose.Cells background image extraction | worksheet background image to Base64 | load workbook from stream Aspose.Cells | convert image bytes to Base64 C# | Excel background picture encoding .NET

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // The example loads an Excel workbook from a stream, accesses the first worksheet, extracts its BackgroundImage byte array (if present), and returns the image encoded as a Base64 string; null is returned when no background image exists or an error occurs.
    public class BackgroundImageHelper
    {
        /// <param name="excelStream">Stream containing the Excel file.</param>
        /// <returns>Base64 representation of the background image or null.</returns>
        public string GetBackgroundImageBase64(Stream excelStream)
        {
            if (excelStream == null) throw new ArgumentNullException(nameof(excelStream));

            try
            {
                // Load the workbook from the provided stream
                var workbook = new Workbook(excelStream);

                // Access the first worksheet (adjust index if needed)
                var sheet = workbook.Worksheets[0];

                // Retrieve the background image as a byte array; may be null if not set
                byte[] bgImage = sheet.BackgroundImage;

                if (bgImage == null || bgImage.Length == 0)
                {
                    // No background image present
                    return null;
                }

                // Encode the byte array to a Base64 string
                return Convert.ToBase64String(bgImage);
            }
            catch (Exception ex)
            {
                // Log the error and return null to indicate failure
                Console.Error.WriteLine($"Error extracting background image: {ex.Message}");
                return null;
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Expect the first argument to be the path to an Excel file
            if (args.Length == 0)
            {
                Console.WriteLine("Please provide the path to an Excel file as an argument.");
                return;
            }

            string filePath = args[0];

            // Prevent FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    var helper = new BackgroundImageHelper();
                    string base64 = helper.GetBackgroundImageBase64(fs);

                    if (!string.IsNullOrEmpty(base64))
                    {
                        Console.WriteLine("Background image Base64:");
                        Console.WriteLine(base64);
                    }
                    else
                    {
                        Console.WriteLine("No background image found in the workbook.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
