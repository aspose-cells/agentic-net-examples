// Title: Set the first worksheet row as a repeated print title on each printed page using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that creates a workbook, adds header values to row 1, fills sample data, sets PageSetup.PrintTitleRows to "$1:$1", and saves the file as an .xlsx document. | Generate a .NET program that configures a worksheet's page setup to repeat the top row on every printed page by using the PrintTitleRows property of Aspose.Cells. | Provide a concise C# snippet demonstrating how to apply repeat‑row print titles in Aspose.Cells and export the workbook.
// Common Searches: Aspose.Cells C# how to repeat header row on each printed page | Set PrintTitleRows property in Aspose.Cells .NET example | Repeat first row as print titles in Excel using Aspose.Cells API | C# code to configure page setup for repeating rows with Aspose.Cells | Aspose.Cells print titles for large worksheets
// Tags: Aspose.Cells PageSetup PrintTitleRows | C# repeat header row on printed Excel pages | Aspose.Cells set print titles for worksheet | Excel repeat rows per page Aspose.Cells .NET | Aspose.Cells workbook page setup configuration | C# Aspose.Cells print title rows example

using Aspose.Cells;
using System;

// The program creates a new workbook, writes column headings in the first row, populates rows 2‑100 with sample data, configures PageSetup.PrintTitleRows to "$1:$1" so the first row repeats on every printed page, and saves the workbook as PrintTitlesExample.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data: column headings in the first row
        sheet.Cells["A1"].PutValue("Header1");
        sheet.Cells["B1"].PutValue("Header2");
        sheet.Cells["C1"].PutValue("Header3");

        // Fill some rows with data (optional)
        for (int i = 2; i <= 100; i++)
        {
            sheet.Cells[$"A{i}"].PutValue($"Row{i - 1}Col1");
            sheet.Cells[$"B{i}"].PutValue($"Row{i - 1}Col2");
            sheet.Cells[$"C{i}"].PutValue($"Row{i - 1}Col3");
        }

        // Set print titles to repeat the first row on each printed page
        sheet.PageSetup.PrintTitleRows = "$1:$1";

        // Save the workbook
        workbook.Save("PrintTitlesExample.xlsx", SaveFormat.Xlsx);
    }
}
