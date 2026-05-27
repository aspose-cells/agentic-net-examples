using System;
using Aspose.Cells;

namespace AsposeCellsProtectionDemo
{
    class Program
    {
        static void Main()
        {
            // ---------- Create source workbook ----------
            Workbook sourceWb = new Workbook();                     // create new workbook
            Worksheet srcSheet = sourceWb.Worksheets[0];           // get first worksheet
            srcSheet.Name = "DataSheet";
            srcSheet.Cells["A1"].PutValue("Protected Workbook");
            srcSheet.Cells["A2"].PutValue(123);

            // Protect the workbook structure with a password
            sourceWb.Protect(ProtectionType.Structure, "pwd123");

            // Verify protection status
            Console.WriteLine("Source workbook structure protected: " +
                sourceWb.Settings.IsProtected);
            Console.WriteLine("Source workbook protected with password: " +
                sourceWb.IsWorkbookProtectedWithPassword);

            // Save the protected source workbook (optional, for inspection)
            sourceWb.Save("ProtectedSource.xlsx", SaveFormat.Xlsx);

            // ---------- Create destination workbook ----------
            Workbook destWb = new Workbook();                       // empty workbook
            Worksheet destSheet = destWb.Worksheets[0];            // first sheet in destination
            destSheet.Name = "CopyTarget";

            // Attempt to copy the protected worksheet into the destination workbook
            try
            {
                // Copy contents and formats from source sheet to destination sheet
                destSheet.Copy(srcSheet);
                Console.WriteLine("Worksheet copied successfully despite source protection.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error during copy operation: " + ex.Message);
            }

            // Save the destination workbook to verify the result
            destWb.Save("CopyResult.xlsx", SaveFormat.Xlsx);
        }
    }
}