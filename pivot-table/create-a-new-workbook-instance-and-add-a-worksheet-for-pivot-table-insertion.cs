// Title: Create a new Aspose.Cells workbook, add a "PivotData" worksheet, populate sample category/amount data, and save as XLSX using C#
// AI Prompts: Generate C# code that uses Aspose.Cells to instantiate a Workbook, add a worksheet named "PivotData", write the headers "Category" and "Amount", fill a few rows with sample values, and save the file as "PivotWorkbook.xlsx". | Show how to prepare a source sheet for a pivot table in Aspose.Cells by creating a workbook, adding a worksheet, inserting sample data rows, and persisting the workbook in XLSX format.
// Common Searches: how to add a worksheet named PivotData in Aspose.Cells C# | Aspose.Cells create workbook and fill sample data for pivot table | C# example of initializing source sheet for a pivot table using Aspose.Cells | save Aspose.Cells workbook as XLSX after adding data sheet
// Tags: create workbook Aspose.Cells C# | add worksheet PivotData Aspose.Cells | populate sample data Aspose.Cells | save workbook as XLSX Aspose.Cells

using System;
using Aspose.Cells;

namespace PivotTableSetup
{
    // // This program creates a new Aspose.Cells Workbook, adds a worksheet called "PivotData", writes sample category and amount rows, and saves the workbook as PivotWorkbook.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook instance
            Workbook workbook = new Workbook();

            // Add a new worksheet that will hold the source data for the pivot table
            Worksheet dataSheet = workbook.Worksheets.Add("PivotData");

            // (Optional) Add some sample data to the worksheet
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Amount");
            dataSheet.Cells["A2"].PutValue("Food");
            dataSheet.Cells["B2"].PutValue(120);
            dataSheet.Cells["A3"].PutValue("Travel");
            dataSheet.Cells["B3"].PutValue(300);
            dataSheet.Cells["A4"].PutValue("Supplies");
            dataSheet.Cells["B4"].PutValue(150);

            // Save the workbook (optional, demonstrates that the workbook is valid)
            workbook.Save("PivotWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}
