// Title: Enable Page Break Preview Mode in an Excel Worksheet using Aspose.Cells for .NET (C#)
// Description: The sample creates a new Workbook, activates the worksheet's Page Break Preview via the IsPageBreakPreview property, optionally sets the zoom to 100 % for accurate visual feedback, and saves the file as PageBreakPreviewDemo.xlsx.
// Keywords: Aspose.Cells | C# | Page Break Preview | Worksheet.IsPageBreakPreview | Excel pagination preview | set worksheet zoom Aspose.Cells | save workbook | print layout preview
// Common Searches: Aspose.Cells enable page break preview C# | How to set IsPageBreakPreview property in Aspose.Cells | C# show Excel page breaks before printing with Aspose | Set worksheet zoom to 100% using Aspose.Cells | Save Excel file with page break preview enabled
// Developer Intent: Activate Page Break Preview for a worksheet and write the workbook to disk.
// Use Cases: Visualize how an Excel sheet will paginate when printed, allowing layout adjustments before distribution. | Generate reports that open directly in Page Break Preview for client review or quality control. | Combine preview mode with a fixed zoom level to provide a WYSIWYG representation of page boundaries.
// AI Prompts: Provide C# code to enable Page Break Preview for every worksheet in a workbook using Aspose.Cells. | Show how to toggle IsPageBreakPreview based on a runtime flag and then save the workbook. | Explain how to retrieve calculated page break positions after turning on Page Break Preview with Aspose.Cells.

using System;
using Aspose.Cells;

// The sample creates a new Workbook, activates the worksheet's Page Break Preview via the IsPageBreakPreview property, optionally sets the zoom to 100 % for accurate visual feedback, and saves the file as PageBreakPreviewDemo.xlsx.
public class EnablePageBreakPreviewDemo
{
    public static void Main(string[] args)
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Enable Page Break Preview mode
        worksheet.IsPageBreakPreview = true;

        // Set zoom to 100% for clear visualization (optional)
        worksheet.Zoom = 100;

        // Save the workbook to a file
        workbook.Save("PageBreakPreviewDemo.xlsx");
        Console.WriteLine("Workbook saved as PageBreakPreviewDemo.xlsx");
    }
}
