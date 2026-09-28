// Title: How to copy a password‑protected worksheet to another workbook and handle the exception when the password is omitted using Aspose.Cells for .NET
// AI Prompts: Generate C# code that protects a worksheet with a password, then attempts to duplicate that sheet into a new workbook without providing the password, and captures the resulting CellsException. | Show how to log the error message returned by Aspose.Cells when AddCopy fails on a protected sheet because the password is missing. | Provide a complete example that saves the original protected workbook and the destination workbook after the failed copy operation.
// Common Searches: Aspose.Cells .NET copy protected sheet to another workbook without password error | C# exception thrown when adding copy of a password‑protected worksheet using Aspose.Cells | How to handle CellsException for protected worksheet copy in Aspose.Cells | Why does AddCopy fail on a protected worksheet in Aspose.Cells for .NET
// Tags: addcopy protected worksheet Aspose.Cells | worksheet protection password handling C# | cellsexception protected sheet copy | saving source and destination workbooks Aspose.Cells | copy worksheet without password Aspose.Cells

using System;
using Aspose.Cells;

// Demonstrates protecting a worksheet with a password, attempting to copy it to another workbook without the password, catching the CellsException, and saving both the source and destination workbooks.
class WorksheetProtectionDemo
{
    static void Main()
    {
        try
        {
            // Create the source workbook and add some data
            Workbook sourceWb = new Workbook();
            Worksheet sourceWs = sourceWb.Worksheets[0];
            sourceWs.Name = "ProtectedSheet";

            // Fill some cells with sample data
            sourceWs.Cells["A1"].PutValue("ID");
            sourceWs.Cells["B1"].PutValue("Name");
            sourceWs.Cells["A2"].PutValue(1);
            sourceWs.Cells["B2"].PutValue("Alice");
            sourceWs.Cells["A3"].PutValue(2);
            sourceWs.Cells["B3"].PutValue("Bob");

            // Protect the entire worksheet with a password (oldPassword is not required here)
            sourceWs.Protect(ProtectionType.All, "mySecretPassword", string.Empty);

            // Save the protected workbook (optional, just for inspection)
            sourceWb.Save("SourceProtected.xlsx");

            // Create a destination workbook where we will try to copy the protected sheet
            Workbook destWb = new Workbook();

            try
            {
                // Attempt to copy the protected worksheet without providing the password
                // This operation should fail because the sheet is protected
                destWb.Worksheets.AddCopy(sourceWs.Name);
                Console.WriteLine("Worksheet copied successfully (unexpected).");
            }
            catch (CellsException ex)
            {
                // Expected exception: the sheet is protected and cannot be copied without the password
                Console.WriteLine("Failed to copy protected worksheet without password:");
                Console.WriteLine(ex.Message);
            }

            // Save the destination workbook (it will not contain the copied sheet)
            destWb.Save("Destination.xlsx");
        }
        catch (Exception ex)
        {
            // General exception handling for unexpected errors
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
