// Title: Automatically apply a predefined shadow style to every picture added in an Aspose.Cells worksheet (C#)
// AI Prompts: Create a C# helper that inserts an image into a worksheet cell and sets its shadow attributes through reflection. | Update the picture insertion routine to automatically set predefined blur, transparency, offset, rotation, and visibility for each new picture. | Implement runtime detection of the Shadow property and apply all shadow attributes only when the current Aspose.Cells version supports it.
// Common Searches: C# Aspose.Cells add picture with automatic shadow effect | set default shadow parameters for every inserted image in Excel using Aspose.Cells | use reflection to configure picture shadow properties in Aspose.Cells .NET | apply consistent shadow style to all pictures in a workbook with Aspose.Cells C#
// Tags: Aspose.Cells picture shadow configuration | C# reflection based picture shadow setting | auto apply shadow to inserted images | default shadow style for Excel pictures | Aspose.Cells image formatting C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, inserts a picture into a specified cell, and uses reflection to apply a predefined shadow style (blur, transparency, offsets, rotation, scale, visibility) only when the Shadow feature is available in the Aspose.Cells version, then saves the file.
class Program
{
    // Predefined shadow style constants (except Type, which is not supported in this version)
    private const double PredefinedBlur = 5.0;
    private const double PredefinedTransparency = 0.5;
    private const double PredefinedOffsetX = 3.0;
    private const double PredefinedOffsetY = 3.0;
    private const double PredefinedRotation = 0.0;
    private const double PredefinedScaleX = 100.0;
    private const double PredefinedScaleY = 100.0;
    private const bool PredefinedVisible = true;

    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Path to the image file
            string imagePath = "sampleImage.png";

            // Add a picture and apply the predefined shadow style if the image exists
            if (File.Exists(imagePath))
            {
                AddPictureWithShadow(sheet, 2, 2, imagePath);
            }
            else
            {
                Console.WriteLine($"Image file not found: {imagePath}. Skipping picture insertion.");
            }

            // Save the workbook
            string resultPath = "Result.xlsx";
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully as {resultPath}.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Adds a picture to the specified cell and applies the predefined shadow style (if supported)
    private static void AddPictureWithShadow(Worksheet sheet, int row, int column, string imagePath)
    {
        try
        {
            // Add picture anchored to the upper-left corner of the specified cell
            int pictureIndex = sheet.Pictures.Add(row, column, imagePath);
            Picture picture = sheet.Pictures[pictureIndex];

            // Attempt to set shadow properties via reflection (compatible with versions that support Shadow)
            var shadowProp = picture.GetType().GetProperty("Shadow");
            if (shadowProp != null)
            {
                object shadowObj = shadowProp.GetValue(picture);
                if (shadowObj != null)
                {
                    SetShadowProperty(shadowObj, "Blur", PredefinedBlur);
                    SetShadowProperty(shadowObj, "Transparency", PredefinedTransparency);
                    SetShadowProperty(shadowObj, "OffsetX", PredefinedOffsetX);
                    SetShadowProperty(shadowObj, "OffsetY", PredefinedOffsetY);
                    SetShadowProperty(shadowObj, "Rotation", PredefinedRotation);
                    SetShadowProperty(shadowObj, "ScaleX", PredefinedScaleX);
                    SetShadowProperty(shadowObj, "ScaleY", PredefinedScaleY);
                    SetShadowProperty(shadowObj, "Visible", PredefinedVisible);
                }
            }
            else
            {
                Console.WriteLine("Shadow feature is not available in the current Aspose.Cells version.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to add picture with shadow: {ex.Message}");
        }
    }

    // Helper to set a property on the shadow object via reflection
    private static void SetShadowProperty(object shadowObj, string propertyName, object value)
    {
        var prop = shadowObj.GetType().GetProperty(propertyName);
        if (prop != null && prop.CanWrite)
        {
            prop.SetValue(shadowObj, Convert.ChangeType(value, prop.PropertyType));
        }
    }
}
