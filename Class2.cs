using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;


namespace Ribbon
{
    public class Class2
    {
        [ExcelFunction(Description = "returns array containing the input range values removing any consecutive repeated")]
        public static object[,] RemoveRepeated([ExcelArgument(AllowReference = true)]object range, bool atTop = false)
        {
            ExcelReference theRef = (ExcelReference)range;
            int rows = theRef.RowLast - theRef.RowFirst + 1;
            ExcelReference callerRef = (ExcelReference)XlCall.Excel(XlCall.xlfCaller);

            //create arrays for values, bins and output
            object[,] res = new object[rows, 1];
            string[] vals = new string[rows];

            //transfer all values to value array
            for (int i = 0; i < rows; i++)
            {
                ExcelReference cellRef = new ExcelReference(theRef.RowFirst + i, theRef.RowFirst + i, theRef.ColumnFirst, theRef.ColumnFirst, theRef.SheetId);
                vals[i] = (string)cellRef.GetValue();
            }

            int top = 0;
            int bottom = rows - 1;
            int dir = 1;

            if (atTop)
            {
                top = bottom;
                bottom = 0;
                dir = -1;
            }

            //create output array
            for (int i = top; dir * i < dir * bottom; i += dir)
            {
                if (vals[i] == vals[i + dir])
                    res[i, 0] = "";
                else
                    res[i, 0] = vals[i];
            }

            res[bottom, 0] = vals[bottom];

            return res;
        }


        [ExcelFunction(Description = "returns array containing the input range values removing any consecutive repeated")]
        public static object[,] SubTotals([ExcelArgument(AllowReference = true)]object rangeKeys, [ExcelArgument(AllowReference = true)]object rangeVals, bool atTop = false)
        {
            ExcelReference theRef = (ExcelReference)rangeKeys;
            ExcelReference valsRef = (ExcelReference)rangeVals;
            int rows = theRef.RowLast - theRef.RowFirst + 1;
            ExcelReference callerRef = (ExcelReference)XlCall.Excel(XlCall.xlfCaller);

            //create arrays for values, bins and output
            object[,] res = new object[rows, 1];
            string[] keys = new string[rows];
            double[] vals = new double[rows];

            //transfer all values to value array
            for (int i = 0; i < rows; i++)
            {
                ExcelReference cellRef = new ExcelReference(theRef.RowFirst + i, theRef.RowFirst + i, theRef.ColumnFirst, theRef.ColumnFirst, theRef.SheetId);
                keys[i] = (string)cellRef.GetValue();
                cellRef = new ExcelReference(valsRef.RowFirst + i, valsRef.RowFirst + i, valsRef.ColumnFirst, valsRef.ColumnFirst, valsRef.SheetId);
                vals[i] = (double)cellRef.GetValue();
            }

            int top = 0;
            int bottom = rows - 1;
            int dir = 1;
            double sum = 0;

            if (atTop)
            {
                top = bottom;
                bottom = 0;
                dir = -1;
            }

            //create output array
            for (int i = top; dir * i < dir * bottom; i += dir)
            {
                if (keys[i] == keys[i + dir])
                {
                    res[i, 0] = "";
                    sum += vals[i];
                }
                else
                {
                    res[i, 0] = sum + vals[i];
                    sum = 0;
                }
            }

            res[bottom, 0] = sum + vals[bottom];

            return res;
        }
    }
}

