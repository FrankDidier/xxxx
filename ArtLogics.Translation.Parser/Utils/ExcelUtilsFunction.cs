using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ArtLogics.Translation.Parser.Utils
{
    public static class ExcelUtilsFunction
    {
        public static int GetColumnNumber(int row, string columnName, ExcelWorksheet worksheet)
        {
            try
            {
                int column;
                for (column = 1; (string)worksheet.Cells[row, column].Value != columnName; column++) ;

                return column;
            } catch {
                return 0;
            }
        }

        public static int GetColumnNumber(int v, object canFunctionNoColName, ExcelWorksheet worksheet)
        {
            throw new NotImplementedException();
        }
    }
}
