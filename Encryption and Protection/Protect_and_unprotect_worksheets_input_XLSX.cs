using System;
using Aspose.Cells;

public class WorksheetProtectionDemo
{
    public static void Run()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Protect the worksheet with a password and all protection types
        sheet.Protect(ProtectionType.All, "pwd123", null);
        Console.WriteLine("Worksheet protected: " + sheet.IsProtected);

        // Save the workbook with protection applied
        workbook.Save("protected.xlsx");

        // Load the protected workbook
        Workbook protectedWb = new Workbook("protected.xlsx");
        Worksheet protectedSheet = protectedWb.Worksheets[0];

        // Unprotect the worksheet using the correct password
        protectedSheet.Unprotect("pwd123");
        Console.WriteLine("Worksheet unprotected: " + !protectedSheet.IsProtected);

        // Save the workbook after unprotecting
        protectedWb.Save("unprotected.xlsx");
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        WorksheetProtectionDemo.Run();
    }
}