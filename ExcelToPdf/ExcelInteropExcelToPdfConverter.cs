using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

namespace ExcelToPdf
{
    public class ExcelInteropExcelToPdfConverter
    {
        public void ConvertToPdf(IEnumerable<string> excelFilesPathToConvert)
        {
            using (var excelApplication = new ExcelApplicationWrapper())
            {
                foreach (var excelFilePath in excelFilesPathToConvert)
                {
                    var thisFileWorkbook = excelApplication.ExcelApplication.Workbooks.Open(excelFilePath);
                    string newPdfFilePath = Path.Combine(
                        Path.GetDirectoryName(excelFilePath),
                        $"{Path.GetFileNameWithoutExtension(excelFilePath)}.pdf");

                    thisFileWorkbook.ExportAsFixedFormat(
                        Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF,
                        newPdfFilePath);

                    thisFileWorkbook.Close(false, excelFilePath);
                    Marshal.ReleaseComObject(thisFileWorkbook);
                }
            }
        }
    }
}