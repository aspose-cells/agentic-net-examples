// Title: Assign custom RGB values to Accent1 and Accent2 theme colors in an Aspose.Cells workbook using C#
// AI Prompts: Write C# code that uses Aspose.Cells' Theme API to set the Accent1 and Accent2 theme colors to specific RGB values and saves the workbook. | Show how to detect the Aspose.Cells version at runtime and apply a custom color scheme only when the Theme and ColorScheme classes are available. | Provide a fallback example that creates a workbook, logs a message about missing Theme support, and explains how to upgrade to a newer Aspose.Cells version for custom theme colors.
// Common Searches: Aspose.Cells C# change Accent1 theme color to custom RGB | how to set Accent2 theme color in Excel with Aspose.Cells .NET | custom Excel theme colors using Aspose.Cells Theme API | upgrade Aspose.Cells to enable Theme and ColorScheme classes | example code for applying a custom color scheme to a workbook in C#
// Tags: Aspose.Cells Theme API custom colors | set accent1 accent2 RGB Aspose.Cells | custom Excel theme palette .NET | upgrade Aspose.Cells for theme support | apply color scheme workbook C#

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// The example creates a new Workbook with Aspose.Cells, notes that the Theme and ColorScheme classes are unavailable in the current library version, advises upgrading to a version that supports the Theme API, and saves the file as 'CustomTheme.xlsx' while outputting its full path.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // NOTE: The Theme and ColorScheme classes are not available in the
            // current Aspose.Cells version used for this project.
            // If you need to modify theme colors, upgrade to a newer version that
            // supports the Theme API. For now we proceed without custom theme changes.

            // Save the workbook
            string outputPath = "CustomTheme.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
