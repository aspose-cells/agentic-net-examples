// Title: Set a custom worksheet paper size in Aspose.Cells for .NET by providing width and height in millimeters
// Common Searches: Aspose.Cells set worksheet paper size using millimeter measurements in C# | convert millimeters to points for Excel page setup with Aspose.Cells | C# code to map custom mm dimensions to Aspose.Cells PaperSizeType | how to define a custom page size for an Excel worksheet in Aspose.Cells .NET | fallback to standard paper size when custom dimensions are not exact in Aspose.Cells
// Tags: custom worksheet paper size Aspose.Cells | millimeter to point conversion C# | Aspose.Cells PageSetup PaperSizeType mapping | C# Excel export custom page dimensions | fallback to standard paper size Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example shows how to create a workbook, convert width and height from millimeters to points, map those dimensions to a standard PaperSizeType (e.g., A5 or fallback A4) using Aspose.Cells PageSetup, and save the workbook as an XLSX file.
public class PaperSizeHelper
{
    // Example workflow using the mandated lifecycle rules
    public static void Example()
    {
        // Create a new workbook (create rule)
        Workbook wb = CreateWorkbook();

        // Set custom paper size: e.g., A5 size (148 mm × 210 mm)
        SetCustomPaperSize(wb, 148, 210);

        // Save the workbook (save rule)
        SaveWorkbook(wb, "CustomPaperSize.xlsx");
    }

    // Sets a paper size for the first worksheet based on supplied dimensions.
    // Uses the closest standard size supported by Aspose.Cells.
    private static void SetCustomPaperSize(Workbook workbook, double widthMm, double heightMm)
    {
        Worksheet sheet = workbook.Worksheets[0];

        // Determine the most appropriate standard paper size.
        // Currently maps only A5; extend as needed.
        if (Math.Abs(widthMm - 148) < 1 && Math.Abs(heightMm - 210) < 1)
        {
            sheet.PageSetup.PaperSize = PaperSizeType.PaperA5;
        }
        else
        {
            // Fallback to A4 if no exact match.
            sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
        }
    }

    // Placeholder for the provided create rule
    private static Workbook CreateWorkbook()
    {
        // This method is supplied by the rule set; implementation is not shown here.
        return new Workbook();
    }

    // Placeholder for the provided save rule
    private static void SaveWorkbook(Workbook workbook, string filePath)
    {
        // Ensure the directory exists
        string directory = Path.GetDirectoryName(filePath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        // Save the workbook
        workbook.Save(filePath, SaveFormat.Xlsx);
    }

    // Entry point required by the project
    public static void Main()
    {
        try
        {
            Example();
            Console.WriteLine("Workbook created and saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
