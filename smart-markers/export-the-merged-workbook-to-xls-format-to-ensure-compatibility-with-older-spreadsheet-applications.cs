// Title: Save a merged workbook as Excel 97‑2003 XLS using Aspose.Cells for .NET (C#)
// Common Searches: how to export a merged workbook to XLS with Aspose.Cells in C# | Aspose.Cells save merged workbook as Excel 97-2003 file | C# convert merged XLSX workbook to XLS using Aspose.Cells | Workbook.Save to Excel97To2003 example in .NET | programmatically save workbook in legacy Excel format with Aspose
// Tags: Workbook.Save Excel97To2003 API | Aspose.Cells XLS compatibility mode | C# convert XLSX to XLS programmatically | export workbook to legacy Excel format | save merged workbook as XLS using Aspose

using Aspose.Cells;
using System;

// The example loads an existing merged workbook (MergedWorkbook.xlsx) and saves it as MergedWorkbook.xls by calling Workbook.Save with the SaveFormat.Excel97To2003 option, providing compatibility with Excel 97‑2003 applications.
class Program
{
    static void Main()
    {
        // Load the merged workbook (assumed to be already created/merged)
        Workbook workbook = new Workbook("MergedWorkbook.xlsx");

        // Export the workbook to the older Excel 97-2003 format (XLS)
        workbook.Save("MergedWorkbook.xls", SaveFormat.Excel97To2003);
    }
}
