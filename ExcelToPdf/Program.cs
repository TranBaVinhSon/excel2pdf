using System;
using System.IO;

namespace ExcelToPdf
{
    internal static class Program
    {
        private static void Main(string[] args)
        {
            string directoryWithExcelFiles;
            if (args.Length == 0)
            {
                // If no directory path is passed as argument, consider the current process directory
                directoryWithExcelFiles = Directory.GetCurrentDirectory();
            }
            else
            {
                directoryWithExcelFiles = Path.GetFullPath(args[0]);
            }

            var excelFilesToConvert = Directory.EnumerateFiles(directoryWithExcelFiles, "*.xls");
            var excelInteropExcelToPdfConverter = new ExcelInteropExcelToPdfConverter();

            try
            {
                excelInteropExcelToPdfConverter.ConvertToPdf(excelFilesToConvert);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Something went wrong: {ex.Message}");
                Environment.ExitCode = -1;
                return;
            }

            Console.WriteLine("Operation completed");
        }
    }
}