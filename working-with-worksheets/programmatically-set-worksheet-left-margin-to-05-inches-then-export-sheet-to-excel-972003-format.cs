// Title: Programmatically set a worksheet's left margin to 0.5 inches and save it as an Excel 97‑2003 (.xls) file with Aspose.Cells for .NET (C#)
// AI Prompts: Create a new workbook, adjust the first worksheet's left page margin to 0.5 inches, and export the file in Excel 97‑2003 (.xls) format using Aspose.Cells in C#. | Use Aspose.Cells to set the left margin of a worksheet to half an inch and then save the workbook as an .xls file compatible with Excel 97‑2003.
// Common Searches: C# Aspose.Cells how to change worksheet left margin in inches | Save Aspose.Cells workbook as Excel 97-2003 .xls with custom page margins | Set page setup margins programmatically with Aspose.Cells for .NET | Export worksheet to .xls while preserving margin settings using Aspose.Cells | Aspose.Cells left margin 0.5 inches example code
// Tags: set left margin Aspose.Cells C# | save workbook Excel97To2003 format | page setup margins Aspose.Cells | export worksheet to .xls with custom margins | configure worksheet page layout programmatically

using Aspose.Cells;
using System;

// Creates a new workbook, sets the first worksheet's left page margin to 0.5 inches, and saves the file as an Excel 97‑2003 (.xls) workbook using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set the left margin to 0.5 inches
        sheet.PageSetup.LeftMargin = 0.5;

        // Export the workbook to Excel 97‑2003 format (.xls)
        workbook.Save("output.xls", SaveFormat.Excel97To2003);
    }
}
