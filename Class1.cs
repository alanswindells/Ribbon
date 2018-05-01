using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
//using Excel = Microsoft.Office.Interop.Excel;
using System.Security.Cryptography;


namespace Ribbon
{
    public class ExelUtils
    {
        [ExcelFunction(Description = "Strip given characters from a string")]
        public static string StripAll(string target, string match)
        {
            string stripped = "";

            for (int i = 0; i < target.Length; i++)
            {
                int pos = match.IndexOf(target[i]);
                if (pos < 0) stripped += target[i];
            }
            return stripped;
        }

        [ExcelFunction(Description = "Find first matching character from given string")]
        public static int FindOneOf(string target, string match)
        {
            int pos;

            for (int i = 0; i < match.Length; i++)
            {
                pos = target.IndexOf(match[i]);
                if (pos >= 0) return pos + 1;
            }

            return -1;
        }

        [ExcelFunction(Description = "Number of occurrances of a string within target string")]
        public static int NumOccurances(string target, string match)
        {
            int pos = -1;
            int n = 0;

            for (; (pos = target.IndexOf(match, pos + 1)) >= 0; n++) ;

            return n;
        }

        [ExcelFunction(Description = "Strip all characters except those give from a string")]
        public static string WithOnly(string target, string match)
        {
            string stripped = "";

            for (int i = 0; i < target.Length; i++)
            {
                int pos = match.IndexOf(target[i]);
                if (pos >= 0) stripped += target[i];
            }
            return stripped;
        }

        [ExcelFunction(Description = "Strip all characters except alphabetic and those give from a string")]
        public static string StripNonAlpha(string target, string except = "")
        {
            string stripped = "";

            for (int i = 0; i < target.Length; i++)
            {
                if (char.IsLetter(target[i]) && IsASCII(target[i].ToString()))
                {
                    stripped += target[i];
                }
                else
                {
                    if (except.IndexOf(target[i]) >= 0) stripped += target[i];
                }
            }
            return stripped;
        }

        [ExcelFunction(Description = "Strip all characters except alphanumeric and those give from a string")]
        public static string StripNonAlphaNum(string target, string except = "")
        {
            string stripped = "";

            for (int i = 0; i < target.Length; i++)
            {
                if (char.IsLetterOrDigit(target[i]) && IsASCII(target[i].ToString()))
                {
                    stripped += target[i];
                }
                else
                {
                    if (except.IndexOf(target[i]) >= 0) stripped += target[i];
                }
            }
            return stripped;
        }


        [ExcelFunction(Description = "Strip all alphabetic characters and the characters given")]
        public static string StripAlpha(string target, string and = "")
        {
            string stripped = "";

            for (int i = 0; i < target.Length; i++) if ((!char.IsLetter(target[i]) || !IsASCII(target[i].ToString())) && (and.IndexOf(target[i]) < 0)) stripped += target[i];
            return stripped;
        }


        [ExcelFunction(Description = "Strip all alphanumeric characters and the characters given")]
        public static string StripAlphaNum(string target, string and = "")
        {
            string stripped = "";

            for (int i = 0; i < target.Length; i++) if ((!char.IsLetterOrDigit(target[i]) || !IsASCII(target[i].ToString())) && (and.IndexOf(target[i]) < 0)) stripped += target[i];
            return stripped;
        }

        [ExcelFunction(Description = "")]
        public static bool IsASCII(string value)
        {
            // ASCII encoding replaces non-ascii with question marks, so we use UTF8 to see if multi-byte sequences are there
            return Encoding.UTF8.GetByteCount(value) == value.Length;
        }

        [ExcelFunction(Description = "Strip all characters except alphanumeric and those give from a string")]
        public static bool IsDiacriticCharacter(string target)
        {
            return char.IsLetter(target[0]) && !IsASCII(target);
        }

        [ExcelFunction(Description = "Strip all characters except alphanumeric and those give from a string")]
        public static bool HasDiacriticCharacters(string target)
        {
            for (int i = 0; i < target.Length; i++) if (IsDiacriticCharacter(target[i].ToString())) return true;
            return false;
        }


        [ExcelFunction(Description = "")]
        public static bool IsSwiftX(string value)
        {
            return StripAlphaNum(value, ": '()+,-./") == "";
        }

        [ExcelFunction(Description = "")]
        public static string DistinctChars(string value)
        {
            return new String(value.Distinct().ToArray());
        }

        [ExcelFunction(Description = "")]
        public static string SortChars(string value, bool dedup = false)
        {
            return dedup ? String.Concat(value.Distinct().OrderBy(c => c)) : String.Concat(value.OrderBy(c => c));
        }

        [ExcelFunction(Description = "returns range containing the input range values scaled by the total, i.e. each value as a % of the sum")]
        public static string Combine([ExcelArgument(AllowReference = true)]object range, string separator = "", bool skipBlanks = false)
        {
            ExcelReference theRef = (ExcelReference)range;
            int rows = theRef.RowLast - theRef.RowFirst + 1;
            int cols = theRef.ColumnLast - theRef.ColumnFirst + 1;
            string res = "";

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    ExcelReference cellRef = new ExcelReference(theRef.RowFirst + i, theRef.RowFirst + i, theRef.ColumnFirst + j, theRef.ColumnFirst + j, theRef.SheetId);
                    object val = cellRef.GetValue();
                    if (val is ExcelDna.Integration.ExcelEmpty)
                    {
                        if (!skipBlanks) res = res + separator;
                    }
                    else
                    {
                        res = res + cellRef.GetValue().ToString() + separator;
                    }
                }
            }

            return res.Substring(0, res.Length - separator.Length);
        }

        [ExcelFunction(Description = "Sorts the given vector")]
        public static double[] SortVector(double[] vector)
        {
            Array.Sort(vector);
            return vector;
        }

        [ExcelFunction(Description = "")]
        public static int Levenshtein(string string1, string string2) //O(n*m)
        {
            var string1Length = string1.Length;
            var string2Length = string2.Length;

            var matrix = new int[string1Length + 1, string2Length + 1];

            // First calculation, if one entry is empty return full length
            if (string1Length == 0)
                return string2Length;

            if (string2Length == 0)
                return string1Length;

            // Initialization of matrix with row size string1Length and columns size string2Length
            for (var i = 0; i <= string1Length; matrix[i, 0] = i++) { }
            for (var j = 0; j <= string2Length; matrix[0, j] = j++) { }

            // Calculate rows and collumns distances
            for (var i = 1; i <= string1Length; i++)
            {
                for (var j = 1; j <= string2Length; j++)
                {
                    var cost = (string2[j - 1] == string1[i - 1]) ? 0 : 1;

                    matrix[i, j] = Math.Min(
                        Math.Min(matrix[i - 1, j] + 1, matrix[i, j - 1] + 1),
                        matrix[i - 1, j - 1] + cost);
                }
            }
            // return result
            return matrix[string1Length, string2Length];
        }

        [ExcelFunction(Description = "")]
        public static int DamerauLevenshtein(string string1, string string2)
        {
            var bounds = new { Height = string1.Length + 1, Width = string2.Length + 1 };

            int[,] matrix = new int[bounds.Height, bounds.Width];

            for (int height = 0; height < bounds.Height; height++) { matrix[height, 0] = height; };
            for (int width = 0; width < bounds.Width; width++) { matrix[0, width] = width; };

            for (int height = 1; height < bounds.Height; height++)
            {
                for (int width = 1; width < bounds.Width; width++)
                {
                    int cost = (string1[height - 1] == string2[width - 1]) ? 0 : 1;
                    int insertion = matrix[height, width - 1] + 1;
                    int deletion = matrix[height - 1, width] + 1;
                    int substitution = matrix[height - 1, width - 1] + cost;

                    int distance = Math.Min(insertion, Math.Min(deletion, substitution));

                    if (height > 1 && width > 1 && string1[height - 1] == string2[width - 2] && string1[height - 2] == string2[width - 1])
                    {
                        distance = Math.Min(distance, matrix[height - 2, width - 2] + cost);
                    }

                    matrix[height, width] = distance;
                }
            }

            return matrix[bounds.Height - 1, bounds.Width - 1];
        }

        [ExcelFunction(Description = "Returns the position of the first different character between the 2 strings")]
        public static int ChangePosition(string string1, string string2)
        {
            if (string1.Length > string2.Length) return ChangePosition(string2, string1);

            for (int i = 0; i < string1.Length; i++)
            {
                if (string1[i] != string2[i]) return i + 1;
            }

            return string2.Length > string1.Length ? string1.Length + 1 : 0;
        }

        [ExcelFunction(Description = "Returns the first different character in the first string")]
        public static string ChangedCharFrom(string string1, string string2)
        {
            int n = ChangePosition(string1, string2) - 1;

            return n < string1.Length ? string1[n].ToString() : "";
        }

        [ExcelFunction(Description = "Returns the first different character in the second string")]
        public static string ChangedCharTo(string string1, string string2)
        {
            int n = ChangePosition(string1, string2) - 1;

            return n < string2.Length ? string2[n].ToString() : "";
        }

        [ExcelFunction(Description = "Returns the position of the first different word")]
        public static int ChangedWordPosition(string string1, string string2)
        {
            int n = ChangePosition(string1, string2) - 1;

            return n >= 0 ? NumOccurances(string1.Substring(0, n), " ") + 1 : 0;
        }

        [ExcelFunction(Description = "Returns the first different word")]
        public static string ChangedWord(string string1, string string2)
        {
            return WordNum(string1, ChangedWordPosition(string1, string2));
        }

        [ExcelFunction(Description = "Returns word at given position in string")]
        public static string WordNum(string target, int n)
        {
            string[] words = target.Split(' ');
            return n <= words.Length ? words[n - 1] : "";
        }


        //returns range containing freq distribution of input range. Optional argument to set the format of the bin labels
        //if return range has two columns then the first column will contain strings describing the bins using the format string
        [ExcelFunction(Description = "Histogram of selected range with given format for labels")]
        public static object[,] Histogram([ExcelArgument(AllowReference = true)]object range, string format = "{0} - {1}")
        {
            ExcelReference theRef = (ExcelReference)range;
            int rows = theRef.RowLast - theRef.RowFirst + 1;
            ExcelReference callerRef = (ExcelReference)XlCall.Excel(XlCall.xlfCaller);
            int nBins = callerRef.RowLast - callerRef.RowFirst + 1;
            int cols = callerRef.ColumnLast - callerRef.ColumnFirst + 1;

            //create arrays for values, bins and output
            object[,] res = new object[nBins, cols];
            double[] vals = new double[rows];
            int[] bins = new int[nBins];
            if (format == "") format = "{0} - {1}";

            //transfer all values to value array
            for (int i = 0; i < rows; i++)
            {
                ExcelReference cellRef = new ExcelReference(theRef.RowFirst + i, theRef.RowFirst + i, theRef.ColumnFirst, theRef.ColumnFirst, theRef.SheetId);
                vals[i] = (double)cellRef.GetValue();
            }

            //calculate range and binsize
            double minVal = vals.Min();
            double maxVal = vals.Max();
            double scale = (maxVal - minVal) / (double)nBins;
            double magnitude = Math.Pow(10, Math.Floor(Math.Log10(scale)));
            double binsize = (1 + (int)(scale / magnitude)) * magnitude;
            double low = (int)(minVal / binsize) * binsize;

            if (low + nBins * binsize < maxVal)
            {
                binsize = (2 + (int)(scale / magnitude)) * magnitude;
                low = (int)(minVal / binsize) * binsize;
            }

            //get frequencies
            Array.Clear(bins, 0, nBins);

            for (int i = 0; i < rows; i++)
            {
                int n = (int)((vals[i] - low) / binsize);
                if (n == nBins) n--;    //max val will be just on the point outside of the last range, so bring it back in
                bins[n]++;
            }

            //create output array
            for (int i = 0; i < nBins; i++)
            {
                res[i, cols - 1] = (int)bins[i];
                if (cols > 1) res[i, 0] = string.Format(format, low + binsize * i, low + binsize * (i + 1));    //labels
            }

            return res;
        }

        [ExcelFunction(Description = "returns range containing the input range values scaled by the total, i.e. each value as a % of the sum")]
        public static object[,] PercentageOfTotal([ExcelArgument(AllowReference = true)]object range)
        {
            ExcelReference theRef = (ExcelReference)range;
            int rows = theRef.RowLast - theRef.RowFirst + 1;
            ExcelReference callerRef = (ExcelReference)XlCall.Excel(XlCall.xlfCaller);

            //create arrays for values, bins and output
            object[,] res = new object[rows, 1];
            double[] vals = new double[rows];

            //transfer all values to value array
            for (int i = 0; i < rows; i++)
            {
                ExcelReference cellRef = new ExcelReference(theRef.RowFirst + i, theRef.RowFirst + i, theRef.ColumnFirst, theRef.ColumnFirst, theRef.SheetId);
                vals[i] = (double)cellRef.GetValue();
            }

            //get sum
            double sum = vals.Sum();

            //create output array
            for (int i = 0; i < rows; i++)
            {
                res[i, 0] = vals[i] / sum;
            }

            return res;
        }

        [ExcelFunction(Description = "")]
        public static object[,] Random()
        {
            ExcelReference callerRef = (ExcelReference)XlCall.Excel(XlCall.xlfCaller);
            int rows = callerRef.RowLast - callerRef.RowFirst + 1;
            int cols = callerRef.ColumnLast - callerRef.ColumnFirst + 1;

            //create arrays for values, bins and output
            object[,] res = new object[rows, cols];

            //transfer all values to value array
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    res[i, j] = ((double)NextRand() + 2147483648.0) / 4294967295.0;
                }
            }

            return res;
        }

        private static int NextRand()
        {
            using (RNGCryptoServiceProvider rg = new RNGCryptoServiceProvider())
            {
                byte[] rno = new byte[5];
                rg.GetBytes(rno);
                return BitConverter.ToInt32(rno, 0);
            }
        }

        [ExcelFunction(Description = "Splits a phrase into words using the specified delimiters")]
        public static object[,] SplitIntoWords(string phrase, string delimiters = " ", bool removeBlanks = false)
        {
            ExcelReference callerRef = (ExcelReference)XlCall.Excel(XlCall.xlfCaller);
            int cols = callerRef.ColumnLast - callerRef.ColumnFirst + 1;

            //create arrays for values, bins and output
            object[,] res = new object[1, cols];

            StringSplitOptions opt = removeBlanks ? StringSplitOptions.RemoveEmptyEntries : StringSplitOptions.None;

            string[] words = phrase.Split(delimiters.ToCharArray(), opt);

            //transfer all values to value array
            for (int j = 0; j < cols; j++)
            {
                res[0, j] = "";
                if (j < words.Length) res[0, j] = words[j];
            }

            return res;
        }
    }
}

