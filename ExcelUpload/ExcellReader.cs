using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Web;
using ExcelUpload.Abstract;

namespace ExcelUpload
{
    public class ExcellReader
    {
        public void ReadExcelContent(HttpPostedFileBase File, List<IPTable> data)
        {
            uint rowNum = 0;
            uint colNum = 0;
            try
            {
                Net.SourceForge.Koogra.IWorkbook wb = null;
                string fileExt = Path.GetExtension(File.FileName);

                if (string.IsNullOrEmpty(fileExt))
                {
                    throw new Exception("File extension not found");
                }

                if (fileExt.Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
                {
                    wb = Net.SourceForge.Koogra.WorkbookFactory.GetExcel2007Reader(File.InputStream);
                }
                else if (fileExt.Equals(".xls", StringComparison.OrdinalIgnoreCase))
                {
                    wb = Net.SourceForge.Koogra.WorkbookFactory.GetExcelBIFFReader(File.InputStream);
                }

                for(int i = 0; i< wb.Worksheets.Count; i++)
                { 

                Net.SourceForge.Koogra.IWorksheet ws = wb.Worksheets.GetWorksheetByIndex(i);

                    IPTable table = data.FirstOrDefault(f => f.Name == ws.Name);
                    if (table == default(IPTable)) continue;

                for (uint r = ws.FirstRow; r <= ws.LastRow; ++r)
                {
                        rowNum = r;
                    Net.SourceForge.Koogra.IRow row = ws.Rows.GetRow(r);
                    if (row != null)
                    {
                        for (uint colCount = ws.FirstCol; colCount <= ws.LastCol; ++colCount)
                        {
                            colNum = colCount;
                            string cellData = string.Empty;
                            if (row.GetCell(colCount) != null && row.GetCell(colCount).Value != null)
                            {
                                cellData = row.GetCell(colCount).Value.ToString();
                            }

                            if (r == ws.FirstRow)
                                table.AddHeader(cellData);
                            else
                                table.Flow = cellData;
                        }
                    }
                }
                }

            }
            catch (Exception ex)
            {
                Exception exception = ex;
                exception.Source = string.Format("Error occured on row {0} col {1}", rowNum, colNum);
                throw ex;
            }


        }
    }
}
