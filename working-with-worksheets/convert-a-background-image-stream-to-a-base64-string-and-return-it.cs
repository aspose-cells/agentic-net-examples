// Title: Convert an Excel worksheet's background image to a Base64 string using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells read worksheet background picture and return Base64 | How to extract Excel sheet background image without System.Drawing | Retrieve first worksheet background image as PNG Base64 with Aspose.Cells | Encode worksheet background image to Base64 using reflection in .NET | Load workbook from stream and get background image Base64 in C#
// Tags: worksheet background image extraction Aspose.Cells | background image to Base64 conversion .NET | load workbook from stream Aspose.Cells | reflection based image saving C# | encode Excel background picture PNG

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The GetWorksheetBackgroundImageBase64 method loads an Excel workbook from a stream, accesses the first worksheet, obtains its BackgroundImage, and uses reflection to invoke the Image.Save method with PNG format. The image bytes are written to a MemoryStream, converted to a Base64 string, and returned; null is returned if no background image exists or any step fails.
    public class BackgroundImageHelper
    {
        /// <param name="workbookStream">Stream containing the Excel workbook.</param>
        /// <returns>Base64 string of the background image, or null.</returns>
        public string GetWorksheetBackgroundImageBase64(Stream workbookStream)
        {
            if (workbookStream == null) throw new ArgumentNullException(nameof(workbookStream));

            try
            {
                // Load the workbook from the provided stream.
                Workbook workbook = new Workbook(workbookStream);

                // Access the first worksheet (modify index if needed).
                Worksheet sheet = workbook.Worksheets[0];

                // Retrieve the background image (type is System.Drawing.Image internally).
                object bgImage = sheet.BackgroundImage;

                // If no background image is set, return null.
                if (bgImage == null)
                    return null;

                // Use reflection to avoid a direct dependency on System.Drawing.Common.
                Type imageFormatType = Type.GetType("System.Drawing.Imaging.ImageFormat");
                if (imageFormatType == null)
                    return null; // Unable to locate ImageFormat type.

                // Get the static property ImageFormat.Png.
                PropertyInfo pngProperty = imageFormatType.GetProperty("Png", BindingFlags.Public | BindingFlags.Static);
                object pngFormat = pngProperty?.GetValue(null);
                if (pngFormat == null)
                    return null; // PNG format not available.

                // Locate the Save(Stream, ImageFormat) method on the image object.
                MethodInfo saveMethod = bgImage.GetType().GetMethod("Save", new[] { typeof(Stream), imageFormatType });
                if (saveMethod == null)
                    return null; // Save method not found.

                // Convert the image to a Base64 string.
                using (MemoryStream ms = new MemoryStream())
                {
                    // Invoke Save to write PNG data into the memory stream.
                    saveMethod.Invoke(bgImage, new object[] { ms, pngFormat });
                    byte[] imageBytes = ms.ToArray();
                    return Convert.ToBase64String(imageBytes);
                }
            }
            catch (Exception ex)
            {
                // Log or handle the exception as needed.
                Console.Error.WriteLine($"Error extracting background image: {ex.Message}");
                return null;
            }
        }
    }

    internal class Program
    {
        private static void Main(string[] args)
        {
            // Example usage:
            // args[0] should contain the path to the Excel file.
            if (args.Length == 0)
            {
                Console.WriteLine("Please provide the path to an Excel file as an argument.");
                return;
            }

            string filePath = args[0];

            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    BackgroundImageHelper helper = new BackgroundImageHelper();
                    string base64 = helper.GetWorksheetBackgroundImageBase64(fs);

                    if (base64 != null)
                    {
                        Console.WriteLine("Background image (Base64):");
                        Console.WriteLine(base64);
                    }
                    else
                    {
                        Console.WriteLine("No background image found in the first worksheet.");
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
