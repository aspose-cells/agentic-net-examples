// Title: Delete rows 60‑65 from an Excel worksheet and export the workbook to PDF using Aspose.Cells for .NET
// AI Prompts: Generate C# code that removes rows 60‑65 from the first sheet of an Excel file and saves the result as a PDF with Aspose.Cells. | Show how to delete a range of rows in a workbook, persist the changes as XLSX, and then use ConversionUtility to produce a PDF. | Provide a step‑by‑step example of using Aspose.Cells to delete rows by index and export the modified workbook to PDF.
// Common Searches: Aspose.Cells C# delete rows 60 to 65 and export to PDF | How to remove a specific row range from an Excel file before converting to PDF in .NET | C# code example for deleting rows in an Excel worksheet and saving as PDF using Aspose.Cells | Delete rows 60‑65 in Excel with Aspose.Cells then generate PDF output
// Tags: Aspose.Cells delete rows worksheet | Aspose.Cells convert workbook to PDF | Aspose.Cells SaveFormat Xlsx after modification | C# remove specific rows Excel Aspose | ConversionUtility Excel to PDF Aspose

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsExample
{
    // // Loads an Excel file, deletes rows 60‑65 from the first worksheet, saves the intermediate workbook as XLSX, and converts it to PDF using Aspose.Cells.
    class DeleteRowsAndConvertToPdf
    {
        static void Main()
        {
            // Path to the original Excel file
            string sourceFile = "input.xlsx";

            // Path for the intermediate workbook after row deletion
            string modifiedFile = "modified.xlsx";

            // Path for the final PDF output
            string pdfFile = "output.pdf";

            // Load the workbook (load rule)
            Workbook workbook = new Workbook(sourceFile);

            // Delete rows 60 through 65 (1‑based). 
            // Cells indices are zero‑based, so start at row index 59 and delete 6 rows.
            workbook.Worksheets[0].Cells.DeleteRows(59, 6);

            // Save the modified workbook (save rule)
            workbook.Save(modifiedFile, SaveFormat.Xlsx);

            // Convert the modified Excel file to PDF (conversion rule)
            ConversionUtility.Convert(modifiedFile, pdfFile);

            Console.WriteLine("Rows deleted and PDF saved successfully.");
        }
    }
}
