// Title: Set a worksheet's left margin to 0.5 inches and save the workbook as an Excel 97‑2003 (.xls) file using Aspose.Cells for .NET
// AI Prompts: Apply Aspose.Cells to set the left page margin of the first worksheet to 0.5 inches, then save the workbook as an Excel 97‑2003 (.xls) file. | In C#, modify a worksheet's left margin to half an inch using Aspose.Cells and export the workbook to the legacy .xls format.
// Common Searches: c# aspocells set left page margin to half an inch | how to export workbook to Excel 97-2003 format with custom margins using Aspose.Cells | Aspose.Cells change worksheet margins programmatically .NET | save workbook as .xls with specific page setup in C# | adjust left margin of worksheet before saving as legacy Excel file
// Tags: worksheet left margin inches Aspose.Cells | save workbook as xls Excel97To2003 | page setup margin configuration .NET | export to legacy Excel format C# | custom page margins Aspose.Cells

using Aspose.Cells;
using System;

// // Creates a workbook, sets the first worksheet's left margin to 0.5 inches, and saves it as an Excel 97‑2003 (.xls) file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set the left margin to 0.5 inches
        sheet.PageSetup.LeftMargin = 0.5;

        // Export the workbook to Excel 97‑2003 format (.xls)
        workbook.Save("output.xls", SaveFormat.Excel97To2003);
    }
}
