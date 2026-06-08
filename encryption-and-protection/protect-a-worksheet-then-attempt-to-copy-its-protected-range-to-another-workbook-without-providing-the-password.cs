using System;
using Aspose.Cells;

namespace AsposeCellsProtectedRangeCopyDemo
{
    class Program
    {
        static void Main()
        {
            // ---------- Create source workbook ----------
            Workbook srcWb = new Workbook();
            Worksheet srcSheet = srcWb.Worksheets[0];

            // Fill some data in the range A1:B2 (the range we will protect)
            srcSheet.Cells["A1"].PutValue("Item");
            srcSheet.Cells["B1"].PutValue("Quantity");
            srcSheet.Cells["A2"].PutValue("Apple");
            srcSheet.Cells["B2"].PutValue(10);
            srcSheet.Cells["A3"].PutValue("Banana");
            srcSheet.Cells["B3"].PutValue(20);

            // Add a protected range covering A1:B2 and set a password for the range
            int rangeIndex = srcSheet.AllowEditRanges.Add("MyProtectedRange", 0, 0, 1, 1);
            ProtectedRange protectedRange = srcSheet.AllowEditRanges[rangeIndex];
            protectedRange.Password = "rangePwd";

            // Protect the entire worksheet with a password
            srcSheet.Protect(ProtectionType.All, "sheetPwd", null);

            // Save the source workbook (lifecycle rule)
            srcWb.Save("SourceProtected.xlsx");

            // ---------- Attempt to copy the protected range without providing password ----------
            // Create a destination workbook
            Workbook destWb = new Workbook();
            Worksheet destSheet = destWb.Worksheets[0];

            try
            {
                // Attempt to copy the cells from the protected range.
                // This operation tries to read the source cells, which is allowed,
                // but writing into the destination while the source sheet is protected
                // without unprotecting will raise an exception when we try to modify the source.
                // Here we simulate a copy by reading values and writing them to the destination.
                for (int row = 0; row <= 1; row++) // rows 0 and 1 correspond to A1:B2
                {
                    for (int col = 0; col <= 1; col++) // columns 0 and 1 correspond to A and B
                    {
                        // Read value from source (allowed)
                        object val = srcSheet.Cells[row, col].Value;

                        // Attempt to write value to destination (allowed)
                        destSheet.Cells[row, col].PutValue(val);
                    }
                }

                Console.WriteLine("Copy operation completed without providing the password.");
            }
            catch (Exception ex)
            {
                // Expected: an exception because the source worksheet is protected with a password
                Console.WriteLine("Failed to copy protected range without password: " + ex.Message);
            }

            // Save the destination workbook to see the result of the attempted copy
            destWb.Save("DestinationCopyAttempt.xlsx");
        }
    }
}