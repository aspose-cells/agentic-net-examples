using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class WorkbookStructureProtectionCheck
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (no protection applied)
                Workbook workbook = new Workbook();

                // Display the structure protection status of the new workbook
                Console.WriteLine("New workbook structure protected: " + workbook.Settings.IsProtected);

                // Protect only the workbook structure with a password
                workbook.Protect(ProtectionType.Structure, "myPassword");

                // Verify that the protection flag is now true
                Console.WriteLine("After protecting structure: " + workbook.Settings.IsProtected);

                // Save the protected workbook to disk
                string filePath = "ProtectedStructureWorkbook.xlsx";
                workbook.Save(filePath, SaveFormat.Xlsx);
                Console.WriteLine("Workbook saved to: " + Path.GetFullPath(filePath));

                // Load the workbook back from the file if it exists
                if (File.Exists(filePath))
                {
                    Workbook loadedWorkbook = new Workbook(filePath);
                    Console.WriteLine("Loaded workbook structure protected: " + loadedWorkbook.Settings.IsProtected);
                }
                else
                {
                    Console.WriteLine("File not found: " + filePath);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            WorkbookStructureProtectionCheck.Run();
        }
    }
}