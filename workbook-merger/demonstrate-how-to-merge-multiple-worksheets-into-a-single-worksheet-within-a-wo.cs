using System;
using Aspose.Cells;

namespace AsposeCellsMergeWorksheetsDemo
{
    class Program
    {
        static void Main()
        {
            // -----------------------------
            // Create first source workbook
            // -----------------------------
            Workbook sourceWb1 = new Workbook();                     // create a new workbook
            Worksheet srcSheet1 = sourceWb1.Worksheets[0];           // get the default sheet
            srcSheet1.Name = "SourceSheet1";
            srcSheet1.Cells["A1"].PutValue("Data from Sheet 1");    // add some data

            // -----------------------------
            // Create second source workbook
            // -----------------------------
            Workbook sourceWb2 = new Workbook();
            Worksheet srcSheet2 = sourceWb2.Worksheets[0];
            srcSheet2.Name = "SourceSheet2";
            srcSheet2.Cells["A1"].PutValue("Data from Sheet 2");

            // -----------------------------
            // Create third source workbook
            // -----------------------------
            Workbook sourceWb3 = new Workbook();
            Worksheet srcSheet3 = sourceWb3.Worksheets[0];
            srcSheet3.Name = "SourceSheet3";
            srcSheet3.Cells["A1"].PutValue("Data from Sheet 3");

            // -------------------------------------------------
            // Destination workbook that will hold all worksheets
            // -------------------------------------------------
            Workbook destWorkbook = new Workbook();                  // empty workbook
            Worksheet destSheet = destWorkbook.Worksheets[0];
            destSheet.Name = "MergedSheets";

            // -------------------------------------------------
            // Combine the source workbooks into the destination
            // -------------------------------------------------
            // The Combine method merges the entire workbook (all its worksheets)
            // into the current workbook. After each call the source workbook's
            // worksheets are appended to the destination workbook.
            destWorkbook.Combine(sourceWb1);
            destWorkbook.Combine(sourceWb2);
            destWorkbook.Combine(sourceWb3);

            // -------------------------------------------------
            // Optional: Move the original empty sheet to the end
            // -------------------------------------------------
            // The first sheet created by the default constructor is empty.
            // If you prefer to remove it, you can delete it or move it.
            destWorkbook.Worksheets.RemoveAt(0); // remove the initially empty sheet

            // -----------------------------
            // Save the merged workbook
            // -----------------------------
            destWorkbook.Save("MergedWorksheets.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("All worksheets have been merged into a single workbook: MergedWorksheets.xlsx");
        }
    }
}