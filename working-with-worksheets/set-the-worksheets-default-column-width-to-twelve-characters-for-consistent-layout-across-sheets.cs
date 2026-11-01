// Title: Set a worksheet's default column width to 12 characters with Aspose.Cells for .NET (C#)
// AI Prompts: Use the Worksheet.Cells.StandardWidth property to set the default column width to 12 characters in a new C# Aspose.Cells workbook and save it as XLSX. | Create a workbook, adjust the default column width for all columns on the first sheet to 12 characters before adding any data, then export the file. | Modify an existing worksheet's StandardWidth to 12 characters using the Aspose.Cells API in C#.
// Common Searches: Aspose.Cells C# set default column width to 12 characters | How to change the standard column width for a worksheet in Aspose.Cells .NET | C# Aspose.Cells set column width for all columns before populating data | Set worksheet default column width using StandardWidth property Aspose.Cells | Configure uniform column width in a new Excel file with Aspose.Cells for .NET
// Tags: Aspose.Cells StandardWidth property | default column width worksheet .NET | column width 12 characters Aspose.Cells | uniform column width workbook C# | configure worksheet column width Aspose

using System;
using Aspose.Cells;

// The program creates a new workbook, accesses the first worksheet, sets its default column width to 12 characters via the StandardWidth property, and saves the file as output.xlsx.
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

            // Set the default column width to 12 characters
            sheet.Cells.StandardWidth = 12;

            // Save the workbook
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
