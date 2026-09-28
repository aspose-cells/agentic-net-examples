// Title: Check if an Excel workbook's VBA project is digitally signed using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, confirms a VBA project exists, and returns the VbaProject.IsSigned value. | Show how to safely access the IsSigned property after verifying the VBA project presence and output the signature status to the console.
// Common Searches: how to use Aspose.Cells to determine if a VBA macro is signed in C# | C# Aspose.Cells check VbaProject.IsSigned property before reading | detect digital signature of Excel VBA project with Aspose.Cells .NET
// Tags: Aspose.Cells check VBA digital signature | C# VbaProject.IsSigned usage | detect signed VBA macro in Excel workbook | validate VBA project presence with Aspose.Cells | read VBA signature status from .xlsx

using System;
using Aspose.Cells;

// C# example that loads an Excel file using Aspose.Cells, verifies a VBA project is present, reads its IsSigned property, and prints whether the VBA project is digitally signed.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Check if the workbook contains a VBA project
        if (workbook.VbaProject != null)
        {
            // Use the IsSigned property to determine digital signature status
            bool isSigned = workbook.VbaProject.IsSigned;

            // Output the result
            Console.WriteLine("VBA project signed: " + isSigned);
        }
        else
        {
            Console.WriteLine("The workbook does not contain a VBA project.");
        }
    }
}
