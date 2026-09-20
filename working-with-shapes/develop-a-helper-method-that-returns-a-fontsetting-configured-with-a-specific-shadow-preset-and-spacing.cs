// Title: How to build a reusable FontSetting with custom font name, size, and color in Aspose.Cells for .NET
// AI Prompts: Write a C# helper method that constructs an Aspose.Cells FontSetting using a temporary Workbook and sets the Font's Name, Size, and Color. | Generate code that creates a FontSetting, configures its underlying Font object, and includes robust exception handling for Aspose.Cells.
// Common Searches: aspnet create FontSetting with specific font name size color Aspose.Cells | C# example for initializing FontSetting using temporary workbook worksheets | how to set custom font properties on Aspose.Cells FontSetting helper | Aspose.Cells FontSetting configuration tutorial for .NET developers | best practice for creating reusable FontSetting method in Aspose.Cells
// Tags: Aspose.Cells FontSetting creation | C# temporary workbook for FontSetting | Aspose.Cells set font properties programmatically | FontSetting helper method .NET | configure font name size color Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Provides a static FontHelper class with a GetFontSetting method that creates a temporary Workbook, uses its Worksheets collection to instantiate a FontSetting, sets the Font's Name, Size (cast to int), and Color, and returns the configured object while handling any exceptions.
public static class FontHelper
{
    // Returns a FontSetting configured with the specified font properties.
    public static FontSetting GetFontSetting(string fontName, double fontSize, Color fontColor)
    {
        try
        {
            // Create a temporary workbook to obtain a WorksheetCollection required by FontSetting constructor.
            Workbook tempWorkbook = new Workbook();

            // Initialize FontSetting with a dummy character range (0,0) and the worksheet collection.
            FontSetting fontSetting = new FontSetting(0, 0, tempWorkbook.Worksheets);

            // Configure the underlying Font object.
            Aspose.Cells.Font font = fontSetting.Font;
            font.Name = fontName;
            // Cast to int if the Font.Size property expects an integer value.
            font.Size = (int)fontSize;
            font.Color = fontColor;

            return fontSetting;
        }
        catch (Exception ex)
        {
            // Log the exception and return null.
            Console.Error.WriteLine($"Error creating FontSetting: {ex.Message}");
            return null;
        }
    }
}

public class Program
{
    public static void Main()
    {
        try
        {
            // Example usage of FontHelper.
            FontSetting setting = FontHelper.GetFontSetting("Arial", 12, Color.Black);
            if (setting != null)
            {
                Aspose.Cells.Font font = setting.Font;
                Console.WriteLine($"Font: {font.Name}, Size: {font.Size}, Color: {font.Color}");
            }
            else
            {
                Console.WriteLine("Failed to create FontSetting.");
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
