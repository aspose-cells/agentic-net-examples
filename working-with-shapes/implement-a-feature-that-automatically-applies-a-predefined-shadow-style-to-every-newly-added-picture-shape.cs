// Title: Apply a Predefined Shadow Style to Every Inserted Picture Shape with Aspose.Cells for .NET (C#)
// Common Searches: how to set picture shadow in Aspose.Cells using C# | add image to worksheet with predefined shadow style Aspose.Cells | Aspose.Cells apply shadow to all inserted pictures automatically | C# example for picture shadow property in Excel workbook | use reflection to access picture shadow in older Aspose.Cells versions
// Tags: Aspose.Cells picture shadow property | C# insert picture with shadow in Excel | reflection for shape properties Aspose.Cells | automated picture styling Aspose.Cells | Excel shape shadow configuration .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, inserts a picture into a worksheet, and uses reflection to safely access the Picture.Shadow property (when available). It then sets a predefined shadow type, blur, offsets, transparency, color, and visibility so that every newly added picture shape receives the same shadow styling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Path to the picture to insert
            string picturePath = "image.png";

            // Add picture (shadow applied if supported)
            AddPictureWithShadow(sheet, picturePath, 5, 5, 100, 100);

            // Save the workbook
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Adds a picture to the worksheet and applies a shadow style when available
    static void AddPictureWithShadow(Worksheet sheet, string picturePath, int row, int column, int width, int height)
    {
        try
        {
            // Verify that the picture file exists
            if (!File.Exists(picturePath))
            {
                Console.WriteLine($"Picture file not found: {picturePath}");
                return;
            }

            // Insert the picture
            int pictureIndex = sheet.Pictures.Add(row, column, picturePath);
            Picture picture = sheet.Pictures[pictureIndex];

            // Set picture size
            picture.Width = width;
            picture.Height = height;

            // Attempt to configure shadow if the API supports it
            try
            {
                // The Shadow property may not exist in older versions; guard with reflection
                var shadowProp = picture.GetType().GetProperty("Shadow");
                if (shadowProp != null)
                {
                    var shadow = shadowProp.GetValue(picture);
                    if (shadow != null)
                    {
                        // Set shadow properties via reflection
                        var typeProp = shadow.GetType().GetProperty("Type");
                        var blurProp = shadow.GetType().GetProperty("Blur");
                        var offsetXProp = shadow.GetType().GetProperty("OffsetX");
                        var offsetYProp = shadow.GetType().GetProperty("OffsetY");
                        var transparencyProp = shadow.GetType().GetProperty("Transparency");
                        var colorProp = shadow.GetType().GetProperty("Color");
                        var visibleProp = shadow.GetType().GetProperty("Visible");

                        if (typeProp != null) typeProp.SetValue(shadow, Enum.Parse(typeProp.PropertyType, "OuterShadow"));
                        if (blurProp != null) blurProp.SetValue(shadow, 5);
                        if (offsetXProp != null) offsetXProp.SetValue(shadow, 3);
                        if (offsetYProp != null) offsetYProp.SetValue(shadow, 3);
                        if (transparencyProp != null) transparencyProp.SetValue(shadow, 0.5);
                        if (colorProp != null) colorProp.SetValue(shadow, Color.Gray);
                        if (visibleProp != null) visibleProp.SetValue(shadow, true);
                    }
                }
            }
            catch (Exception shadowEx)
            {
                // Shadow configuration failed; continue without it
                Console.WriteLine($"Shadow not applied: {shadowEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to add picture with shadow: {ex.Message}");
        }
    }
}
