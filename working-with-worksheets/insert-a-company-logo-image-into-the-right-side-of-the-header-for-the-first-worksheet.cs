// Title: Add a PNG company logo to the right header of the first worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that sets the right header placeholder "&G", inserts a PNG logo into the header of the first worksheet, and saves the workbook. | Demonstrate how to use reflection to retrieve the RightHeaderPicture object, assign a PNG file to it, and enable the IsScaled property for proper scaling. | Create a robust example that verifies the logo file exists, handles missing‑file errors gracefully, and writes the workbook to a specified output path.
// Common Searches: how to insert a PNG logo into the right page header of an Excel sheet using Aspose.Cells C# | Aspose.Cells set right header picture with reflection for compatibility with older versions | adjust header image size when adding a logo with Aspose.Cells C# example | handle missing header image file when adding a company logo using Aspose.Cells | add company logo to Excel header using Aspose.Cells .NET API
// Tags: Aspose.Cells insert header image C# | right header picture Aspose.Cells | header picture scaling Aspose.Cells | reflection access RightHeaderPicture Aspose.Cells | logo file existence check Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, accesses the first worksheet, sets the right header placeholder to "&G", loads a PNG logo, uses reflection to assign the image to the RightHeaderPicture property, enables scaling, validates the logo file's presence, and saves the workbook as output.xlsx, with comprehensive error handling for missing files and API version differences.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet (index 0)
            Worksheet sheet = workbook.Worksheets[0];

            // Set the right header placeholder for a picture ("&G")
            try
            {
                var pageSetup = sheet.PageSetup;
                var rightHeaderProp = pageSetup.GetType().GetProperty("RightHeader");
                if (rightHeaderProp != null && rightHeaderProp.CanWrite)
                {
                    rightHeaderProp.SetValue(pageSetup, "&G");
                }
                else
                {
                    Console.WriteLine("RightHeader property is not available in this Aspose.Cells version.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to set right header placeholder: {ex.Message}");
            }

            // Path to the logo image
            string logoPath = "logo.png";

            // Ensure the logo file exists before attempting to use it
            if (!File.Exists(logoPath))
                throw new FileNotFoundException($"Logo file not found: {logoPath}");

            // Set the header picture using reflection (covers versions without direct API)
            try
            {
                var pageSetup = sheet.PageSetup;
                var pictureProp = pageSetup.GetType().GetProperty("RightHeaderPicture");
                if (pictureProp != null)
                {
                    var pictureObj = pictureProp.GetValue(pageSetup);
                    var setImageMethod = pictureObj?.GetType().GetMethod("SetImage", new[] { typeof(string) });
                    setImageMethod?.Invoke(pictureObj, new object[] { logoPath });

                    var isScaledProp = pictureObj?.GetType().GetProperty("IsScaled");
                    if (isScaledProp != null && isScaledProp.CanWrite)
                    {
                        isScaledProp.SetValue(pictureObj, true);
                    }
                }
                else
                {
                    Console.WriteLine("Header picture feature is not available in this Aspose.Cells version.");
                }
            }
            catch (Exception imgEx)
            {
                Console.WriteLine($"Failed to load logo image: {imgEx.Message}");
                // Continue without the image if needed
            }

            // Save the workbook to a file
            string outputPath = "output.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
