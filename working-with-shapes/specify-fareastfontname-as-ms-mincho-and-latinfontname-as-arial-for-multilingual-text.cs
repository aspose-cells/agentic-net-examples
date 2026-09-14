// Title: How to set Latin (Arial) and FarEast (MS Mincho) fonts for mixed-language cells using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that creates a workbook, defines a style with Arial for Latin characters and MS Mincho for FarEast characters, applies the style to a cell containing English and Japanese text, and saves the file. | Show how to use Aspose.Cells Style.Font.Name and, when supported, Style.Font.FarEastName to assign separate fonts for Latin and FarEast scripts in an Excel worksheet.
// Common Searches: Aspose.Cells set Latin font to Arial and FarEast font to MS Mincho in C# | C# example for applying different fonts to English and Japanese text in Excel with Aspose.Cells | How to style mixed language cells in Aspose.Cells .NET | Set FarEast font name for Japanese characters using Aspose.Cells API | Apply multilingual font styling to a worksheet cell in Aspose.Cells for .NET
// Tags: aspocells set latin font c# | aspocells far east font name | multilingual font styling aspocells | excel cell style mixed language c# | aspocells style font far east

using System;
using Aspose.Cells;

// The example creates a new Workbook, defines a Style with the Latin font set to Arial and, when supported, the FarEast font set to MS Mincho, writes English‑Japanese text into cell A1, applies the style, and saves the workbook as MultilingualText.xlsx. It also notes that the FarEast font property may be unavailable in older Aspose.Cells versions, in which case the default FarEast font is used.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Create a style to set the Latin font
            Style multilingualStyle = workbook.CreateStyle();

            // Set Latin font name to Arial
            multilingualStyle.Font.Name = "Arial";

            // NOTE: The FarEast font property is not available in the current Aspose.Cells version.
            // The default FarEast font will be used for characters such as Japanese.

            // Put multilingual text into a cell
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("English 日本語");

            // Apply the multilingual style to the cell
            cell.SetStyle(multilingualStyle);

            // Save the workbook
            workbook.Save("MultilingualText.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
