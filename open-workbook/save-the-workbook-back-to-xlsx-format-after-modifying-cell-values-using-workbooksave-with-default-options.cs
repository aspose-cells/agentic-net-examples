// Title: Modify cells A1 and B2 and save the workbook as XLSX using Aspose.Cells for .NET with default options
// AI Prompts: Load an existing XLSX file, set cell A1 to a string and cell B2 to a numeric value, then call Workbook.Save to write the changes back to a new XLSX file using default save settings in C#. | Using Aspose.Cells, update multiple cells in a worksheet and overwrite the original file without providing custom SaveOptions, demonstrating the default save behavior in .NET.
// Common Searches: asp.net how to change cell values and save workbook without specifying SaveOptions | c# Aspose.Cells overwrite existing xlsx after editing cells | default save behavior of Aspose.Cells when saving modified workbook | example code to modify A1 and B2 then save as new xlsx using Aspose.Cells | save workbook to xlsx format with default options after cell updates in C#
// Tags: modify cell values Aspose.Cells C# | save workbook to xlsx default options | overwrite existing Excel file Aspose.Cells | cell update worksheet Aspose.Cells API | default save behavior Aspose.Cells .NET

using Aspose.Cells;

// Loads input.xlsx, changes A1 to "Hello World" and B2 to 12345, then saves the workbook as output.xlsx using Workbook.Save with default settings.
class Program
{
    static void Main()
    {
        // Load an existing workbook (XLSX format)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Modify some cell values
        sheet.Cells["A1"].PutValue("Hello World");
        sheet.Cells["B2"].PutValue(12345);

        // Save the workbook back to XLSX format using default save options
        workbook.Save("output.xlsx");
    }
}
