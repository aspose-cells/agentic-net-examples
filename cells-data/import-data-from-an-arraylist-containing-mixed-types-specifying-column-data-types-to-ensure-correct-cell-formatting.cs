// Title: Import a mixed‑type ArrayList into an Excel worksheet and apply custom date and currency column formats using Aspose.Cells for .NET (C#)
// Common Searches: c# import arraylist with strings dates numbers into excel using aspose.cells | set custom date format for a column after ImportArrayList in Aspose.Cells | apply currency number format to a specific column in an Aspose.Cells workbook (C#) | preserve data types when importing mixed data into Excel with Aspose.Cells | how to style Excel columns programmatically with Aspose.Cells C# | save workbook as xlsx after formatting columns using Aspose.Cells
// Tags: ImportArrayList mixed data Aspose.Cells C# | date column formatting Aspose.Cells | currency column formatting Aspose.Cells | column style application Aspose.Cells | save workbook as xlsx Aspose.Cells | heterogeneous data import Aspose.Cells

using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, imports a two‑dimensional ArrayList containing strings, integers, dates, and decimals into the first worksheet starting at A1, then applies a custom "mm-dd-yyyy" format to the date column and a "$#,##0.00" format to the salary column before saving the file as MixedDataImport.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Prepare mixed type data in an ArrayList (each inner ArrayList represents a row)
            ArrayList data = new ArrayList
            {
                new ArrayList { "John Doe", 30, DateTime.Parse("1993-04-15"), 1234.56 },
                new ArrayList { "Jane Smith", 25, DateTime.Parse("1998-09-22"), 7890.12 },
                new ArrayList { "Bob Johnson", 40, DateTime.Parse("1983-01-05"), 3456.78 }
            };

            // Import the ArrayList data starting at cell A1 (row 0, column 0)
            // The fourth parameter 'importDataOnly' set to false to allow default formatting
            cells.ImportArrayList(data, 0, 0, false);

            // Apply specific number formats for date and currency columns
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "mm-dd-yyyy";
            StyleFlag dateFlag = new StyleFlag { NumberFormat = true };
            cells.Columns[2].ApplyStyle(dateStyle, dateFlag); // Date column

            Style currencyStyle = workbook.CreateStyle();
            currencyStyle.Custom = "$#,##0.00";
            StyleFlag currencyFlag = new StyleFlag { NumberFormat = true };
            cells.Columns[3].ApplyStyle(currencyStyle, currencyFlag); // Salary column

            // Define output file path
            string outputPath = "MixedDataImport.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
