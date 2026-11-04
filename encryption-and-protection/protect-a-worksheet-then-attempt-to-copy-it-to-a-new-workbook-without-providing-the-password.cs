// Title: Copy a password‑protected worksheet to another workbook without supplying the password using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# copy a protected worksheet to a new workbook without entering the password | How to retain worksheet protection when using Worksheets.AddCopy in Aspose.Cells | Copying a password‑protected Excel sheet programmatically with Aspose.Cells .NET API | Duplicate a protected sheet in Aspose.Cells without providing the password
// Tags: Aspose.Cells Worksheets.AddCopy protected sheet | retain worksheet protection flag Aspose.Cells | duplicate Excel sheet without password C# | protected worksheet copy operation Aspose.Cells | C# copy protected worksheet to new workbook

using System;
using Aspose.Cells;

// The example creates a source workbook, protects its first worksheet with a password, and then copies that protected worksheet into a new workbook using Worksheets.AddCopy without providing the password. The copied sheet retains its protection flag, and both workbooks are saved.
class Program
{
    static void Main()
    {
        try
        {
            // Create a source workbook and protect its first worksheet
            Workbook sourceWorkbook = new Workbook();
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            sourceSheet.Name = "ProtectedSheet";

            // Protect the worksheet with a password (oldPassword is not required here, pass empty string)
            sourceSheet.Protect(ProtectionType.All, "MyPassword", string.Empty);

            // (Optional) Save the source workbook
            sourceWorkbook.Save("Source.xlsx");

            // Create a destination workbook
            Workbook destinationWorkbook = new Workbook();

            // Copy the protected worksheet to the new workbook without providing the password
            // The copy operation succeeds; the protection flag is retained on the copied sheet.
            destinationWorkbook.Worksheets.AddCopy(sourceSheet.Name);
            Console.WriteLine("Worksheet copied successfully.");

            // Verify that the copied sheet is still protected
            Worksheet copiedSheet = destinationWorkbook.Worksheets[destinationWorkbook.Worksheets.Count - 1];
            Console.WriteLine("Copied sheet protected: " + copiedSheet.IsProtected);

            // Save the destination workbook
            destinationWorkbook.Save("Destination.xlsx");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
