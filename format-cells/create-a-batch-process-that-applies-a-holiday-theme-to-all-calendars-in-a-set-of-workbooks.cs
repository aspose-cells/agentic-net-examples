// Title: Apply a holiday-themed cell style and header image to all worksheets in multiple Excel workbooks using Aspose.Cells for .NET
// Common Searches: how to batch apply a custom cell style to DateTime cells in many Excel files with Aspose.Cells C# | add a holiday picture to every worksheet while processing multiple workbooks using Aspose.Cells | C# code to iterate through a folder of .xlsx files and style calendar cells with Aspose.Cells | automate holiday theme insertion into Excel workbooks using Aspose.Cells .NET
// Tags: batch apply cell style Aspose.Cells | style DateTime cells Aspose.Cells | insert worksheet picture Aspose.Cells | process multiple .xlsx files C# | apply holiday theme Excel Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

// The example loads each .xlsx file from an input directory, creates a red background with white bold font style, applies this style to every cell containing a DateTime value, inserts a holiday.png image at the top‑left corner of each worksheet, and saves the modified workbooks to an output folder, enabling a holiday theme across a batch of workbooks.
class HolidayThemeBatchProcessor
{
    // Path to the folder containing the workbooks to process
    private const string InputFolder = @"C:\Workbooks\Input";

    // Path to the folder where processed workbooks will be saved
    private const string OutputFolder = @"C:\Workbooks\Output";

    // Path to the holiday image that will be added to each worksheet
    private const string HolidayImagePath = @"C:\Resources\holiday.png";

    static void Main()
    {
        // Ensure output directory exists
        if (!Directory.Exists(OutputFolder))
            Directory.CreateDirectory(OutputFolder);

        // Process each Excel file in the input folder
        foreach (string filePath in Directory.GetFiles(InputFolder, "*.xlsx"))
        {
            // Load workbook (load rule)
            Workbook workbook = new Workbook(filePath);

            // Apply holiday theme to all worksheets
            ApplyHolidayTheme(workbook);

            // Determine output file path
            string fileName = Path.GetFileName(filePath);
            string outputPath = Path.Combine(OutputFolder, fileName);

            // Save workbook (save rule)
            workbook.Save(outputPath);
        }

        Console.WriteLine("Holiday theme applied to all workbooks.");
    }

    private static void ApplyHolidayTheme(Workbook workbook)
    {
        // Create a style that represents the holiday theme (create rule)
        Style holidayStyle = workbook.CreateStyle();
        holidayStyle.ForegroundColor = Color.Red;               // Red background
        holidayStyle.Pattern = BackgroundType.Solid;
        holidayStyle.Font.Color = Color.White;                  // White font
        holidayStyle.Font.IsBold = true;                        // Bold text

        // Iterate through each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Apply the holiday style to all cells containing DateTime values
            foreach (Cell cell in cells)
            {
                if (cell.Value != null && cell.Value is DateTime)
                {
                    cell.SetStyle(holidayStyle);
                }
            }

            // Insert a holiday image at the top-left corner of the sheet (create rule)
            if (File.Exists(HolidayImagePath))
            {
                // Parameters: upper left row, upper left column, image file path
                sheet.Pictures.Add(0, 0, HolidayImagePath);
            }
        }
    }
}
