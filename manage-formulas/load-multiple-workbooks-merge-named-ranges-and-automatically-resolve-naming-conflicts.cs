using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsMergeNamedRanges
{
    class Program
    {
        static void Main()
        {
            // Paths of workbooks to be merged
            string[] workbookFiles = new string[]
            {
                "Workbook1.xlsx",
                "Workbook2.xlsx",
                "Workbook3.xlsx"
            };

            // Load the first workbook as the target workbook
            Workbook targetWorkbook = new Workbook(workbookFiles[0]);

            // Process remaining workbooks
            for (int i = 1; i < workbookFiles.Length; i++)
            {
                // Load source workbook
                Workbook sourceWorkbook = new Workbook(workbookFiles[i]);

                // Merge named ranges from source into target
                MergeNamedRanges(targetWorkbook, sourceWorkbook);
            }

            // Remove any duplicate names that might still exist
            targetWorkbook.Worksheets.Names.RemoveDuplicateNames();

            // Save the merged workbook
            string outputPath = "MergedWorkbook.xlsx";
            targetWorkbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Merged workbook saved to '{outputPath}'.");
        }

        /// <summary>
        /// Merges named ranges from a source workbook into a target workbook.
        /// If a name conflict occurs, the source name is renamed by appending a numeric suffix.
        /// </summary>
        static void MergeNamedRanges(Workbook target, Workbook source)
        {
            // Iterate through all defined names in the source workbook
            foreach (Name srcName in source.Worksheets.Names)
            {
                // Determine a unique name for the target workbook
                string uniqueName = GetUniqueName(target.Worksheets.Names, srcName.Text);

                // Add the (potentially renamed) name to the target workbook
                int index = target.Worksheets.Names.Add(uniqueName);
                Name newName = target.Worksheets.Names[index];

                // Preserve the reference (formula) of the original name
                newName.RefersTo = srcName.RefersTo;
            }
        }

        /// <summary>
        /// Generates a unique name that does not exist in the provided NameCollection.
        /// If the original name already exists, a suffix "_1", "_2", ... is appended.
        /// </summary>
        static string GetUniqueName(NameCollection names, string baseName)
        {
            // Quick check: if the name does not exist, return it unchanged
            if (!NameExists(names, baseName))
                return baseName;

            // Otherwise, append a numeric suffix until a unique name is found
            int suffix = 1;
            string candidate;
            do
            {
                candidate = $"{baseName}_{suffix}";
                suffix++;
            } while (NameExists(names, candidate));

            return candidate;
        }

        /// <summary>
        /// Checks whether a name with the specified text exists in the NameCollection.
        /// </summary>
        static bool NameExists(NameCollection names, string nameText)
        {
            // NameCollection does not expose a direct Contains method,
            // so we iterate through the collection to find a match.
            foreach (Name n in names)
            {
                if (string.Equals(n.Text, nameText, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }
    }
}