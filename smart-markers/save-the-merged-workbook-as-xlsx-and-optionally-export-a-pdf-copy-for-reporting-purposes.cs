// Title: Combine two Excel workbooks and export the merged file as XLSX and PDF with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to merge two existing Excel workbooks and then save the combined workbook in both XLSX and PDF formats. | Generate a .NET example that demonstrates merging two workbooks and subsequently saving the result as an XLSX file and a PDF file using Aspose.Cells.
// Common Searches: c# Aspose.Cells merge two workbooks and save as xlsx | how to export a combined Excel file to pdf with Aspose.Cells | Aspose.Cells save merged workbook in both xlsx and pdf formats | combine multiple Excel files into one using Aspose.Cells .NET | dual format export of merged workbook Aspose.Cells c#
// Tags: Workbook.Combine method C# | save merged workbook as xlsx Aspose.Cells | export merged workbook to pdf Aspose.Cells | dual-format workbook export Aspose.Cells | combine Excel workbooks Aspose.Cells .NET

using System;
using Aspose.Cells;

// The example creates two separate workbooks, adds sample data, merges the second workbook into the first using Aspose.Cells, and then saves the combined workbook as both MergedWorkbook.xlsx and MergedWorkbook.pdf.
class SaveMergedWorkbook
{
    static void Main()
    {
        // Create the first workbook and add sample data
        Workbook firstWorkbook = new Workbook();
        Worksheet firstSheet = firstWorkbook.Worksheets[0];
        firstSheet.Name = "FirstSheet";
        firstSheet.Cells["A1"].PutValue("Data from the first workbook");

        // Create the second workbook and add sample data
        Workbook secondWorkbook = new Workbook();
        Worksheet secondSheet = secondWorkbook.Worksheets[0];
        secondSheet.Name = "SecondSheet";
        secondSheet.Cells["A1"].PutValue("Data from the second workbook");

        // Merge the second workbook into the first one
        firstWorkbook.Combine(secondWorkbook);

        // Save the merged workbook as XLSX
        firstWorkbook.Save("MergedWorkbook.xlsx", SaveFormat.Xlsx);

        // Optionally, also save a PDF copy for reporting
        firstWorkbook.Save("MergedWorkbook.pdf", SaveFormat.Pdf);
    }
}
