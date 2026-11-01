// Title: Set Excel write‑protection author to the logged‑in Windows user and overwrite the workbook with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set write protection author to Environment.UserName | How to change Excel workbook write‑protection author programmatically using Aspose.Cells | Overwrite existing .xlsx after updating write protection author with Aspose.Cells in .NET | Aspose.Cells example assigning current Windows user as write‑protection author | Save workbook after modifying write protection settings using Aspose.Cells C#
// Tags: Aspose.Cells write protection author | C# set write protection author | Aspose.Cells overwrite workbook file | Environment.UserName Excel protection | Aspose.Cells modify workbook settings

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookWriteProtectionDemo
{
    // The sample checks for the presence of an Excel file, creates a new workbook if needed, loads the workbook, assigns the current Windows user name (Environment.UserName) to the write‑protection author via workbook.Settings.WriteProtection.Author, and then saves the workbook back to the original path, handling file‑not‑found and generic exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook file – adjust as needed.
            string filePath = @"C:\Path\To\YourWorkbook.xlsx";

            try
            {
                // Ensure the file exists; create an empty workbook if it does not.
                if (!File.Exists(filePath))
                {
                    Workbook newWorkbook = new Workbook();
                    newWorkbook.Save(filePath);
                }

                // Load the existing workbook.
                Workbook workbook = new Workbook(filePath);

                // Set the write‑protection author to the current Windows user name.
                workbook.Settings.WriteProtection.Author = Environment.UserName;

                // Re‑save the workbook (overwrites the original file).
                workbook.Save(filePath);

                Console.WriteLine("Workbook write‑protection author set successfully.");
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.Error.WriteLine($"File not found: {fnfEx.FileName}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
