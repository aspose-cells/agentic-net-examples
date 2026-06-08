using Aspose.Cells;
using System;
using System.IO;

class WorksheetCopyProtectedDemo
{
    static void Main()
    {
        try
        {
            // Create a source workbook and protect its first worksheet with a password
            Workbook sourceWorkbook = new Workbook();
            Worksheet sourceWorksheet = sourceWorkbook.Worksheets[0];
            sourceWorksheet.Protect(ProtectionType.All, "mySecretPassword", null);
            Console.WriteLine("Source worksheet protected: " + sourceWorksheet.IsProtected);

            // Create a destination workbook where we will try to copy the protected worksheet
            Workbook destinationWorkbook = new Workbook();

            try
            {
                // Attempt to copy the protected worksheet without providing the password.
                // AddCopy expects the worksheet name, not the Worksheet object.
                int copiedIndex = destinationWorkbook.Worksheets.AddCopy(sourceWorksheet.Name);
                Console.WriteLine("Worksheet copied successfully, new index: " + copiedIndex);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error copying protected worksheet without password: " + ex.Message);
            }

            // Save both workbooks for verification
            sourceWorkbook.Save("SourceProtected.xlsx");
            destinationWorkbook.Save("Destination.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Unexpected error: " + ex.Message);
        }
    }
}