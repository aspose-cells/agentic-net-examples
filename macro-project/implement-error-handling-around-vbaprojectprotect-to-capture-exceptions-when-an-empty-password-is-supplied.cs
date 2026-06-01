using System;
using Aspose.Cells;

class VbaProjectProtectDemo
{
    static void Main()
    {
        // Create a new workbook (macro-enabled)
        Workbook workbook = new Workbook();

        // Attempt to protect the VBA project with an empty password
        try
        {
            // islockedForViewing = false, password = empty string
            workbook.VbaProject.Protect(false, "");
        }
        catch (Exception ex)
        {
            // Capture and display the exception caused by an empty password
            Console.WriteLine("Error protecting VBA project: " + ex.Message);
        }

        // Properly protect the VBA project with a valid password
        try
        {
            workbook.VbaProject.Protect(false, "ValidPassword123");
            Console.WriteLine("VBA project protected successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Unexpected error: " + ex.Message);
        }

        // Save the workbook as a macro-enabled file
        workbook.Save("ProtectedDemo.xlsm", SaveFormat.Xlsm);
    }
}