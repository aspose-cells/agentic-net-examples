using System;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaReflectionDemo
{
    class Program
    {
        static void Main()
        {
            // Load a macro-enabled workbook (replace with actual path if needed)
            // If the file does not exist, a new workbook will be created which also has a VBA project.
            Workbook workbook;
            string inputPath = "SampleWithVba.xlsm";
            if (System.IO.File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                // Ensure the workbook has a VBA project by saving as .xlsm and reloading
                workbook.Save("temp.xlsm", SaveFormat.Xlsm);
                workbook = new Workbook("temp.xlsm");
                System.IO.File.Delete("temp.xlsm");
            }

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;
            if (vbaProject == null)
            {
                Console.WriteLine("No VBA project found in the workbook.");
                return;
            }

            // Get the collection of references
            VbaProjectReferenceCollection references = vbaProject.References;
            if (references == null || references.Count == 0)
            {
                Console.WriteLine("No VBA references found.");
                return;
            }

            // Enumerate references using reflection
            foreach (object refObj in references)
            {
                Type refType = refObj.GetType();

                // Get Name property
                PropertyInfo nameProp = refType.GetProperty("Name");
                string name = nameProp?.GetValue(refObj) as string ?? "Unnamed";

                // Get Libid property (contains version information for registered references)
                PropertyInfo libidProp = refType.GetProperty("Libid");
                string libid = libidProp?.GetValue(refObj) as string ?? string.Empty;

                // Attempt to extract version from Libid (format: *\G{...}#<Version>#...)
                string version = "Unknown";
                if (!string.IsNullOrEmpty(libid))
                {
                    string[] parts = libid.Split('#');
                    if (parts.Length >= 2 && !string.IsNullOrWhiteSpace(parts[1]))
                    {
                        version = parts[1];
                    }
                }

                Console.WriteLine($"Reference Name: {name}, Version: {version}");
            }

            // Optionally save the workbook (demonstrates usage of save rule)
            workbook.Save("EnumeratedVbaReferences.xlsm", SaveFormat.Xlsm);
        }
    }
}