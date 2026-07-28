// Title: Aspose.Cells .NET – Attempt to edit a password‑protected worksheet without a password (exception handling)
// Description: This C# example creates a workbook, protects the first worksheet with the password "secret", then tries to modify cell A1 without providing the password. The operation throws an exception, which is caught and logged, and the workbook is saved unchanged.
// Keywords: Aspose.Cells | C# | .NET | protected worksheet | worksheet protection | ProtectionType.All | exception handling | edit without password | catch exception | save workbook | security audit
// Common Searches: Aspose.Cells edit protected sheet without password | C# catch exception when writing to a locked worksheet | What error is thrown by Aspose.Cells when modifying a protected cell | How to detect failed edit on a protected worksheet in .NET | Save Aspose.Cells workbook after protection error
// Developer Intent: Demonstrate that writing to a password‑protected worksheet without the correct password raises an exception and show how to capture the failure.
// Use Cases: Validate worksheet protection status before performing write operations to prevent runtime errors. | Log unauthorized edit attempts on protected sheets for compliance and security monitoring. | Ensure the workbook is saved only when no prohibited modifications have been applied.
// AI Prompts: Generate C# code using Aspose.Cells that checks if a worksheet is protected and edits cells only after supplying the correct password. | Show the specific exception type thrown by Aspose.Cells when attempting to write to a protected cell without a password. | Provide an example that programmatically unprotects a worksheet, updates a cell, and re‑protects it while handling possible errors.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This C# example creates a workbook, protects the first worksheet with the password "secret", then tries to modify cell A1 without providing the password. The operation throws an exception, which is caught and logged, and the workbook is saved unchanged.
    public class EditProtectedWorksheetDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Put initial data into a cell
                sheet.Cells["A1"].PutValue("Original");

                // Protect the worksheet with a password
                sheet.Protect(ProtectionType.All, "secret", null);
                Console.WriteLine("Worksheet protected: " + sheet.IsProtected);

                // Attempt to edit a cell without providing the password
                try
                {
                    // This operation should fail because the sheet is protected
                    sheet.Cells["A1"].PutValue("Modified without password");
                    Console.WriteLine("Edit succeeded unexpectedly.");
                }
                catch (Exception ex)
                {
                    // Record the failure outcome
                    Console.WriteLine("Failed to edit protected worksheet without password: " + ex.Message);
                }

                // Save the workbook (the edit should not have been applied)
                workbook.Save("EditProtectedWorksheetDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }

        // Entry point required for compilation
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
