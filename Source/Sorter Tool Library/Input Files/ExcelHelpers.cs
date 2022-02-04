using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace SorterToolLibrary.SorterTool
{
    public static class ExcelHelpers
    {
        public enum CharType
        {
            Unknown,
            Number,
            Letter,
            Punctuation
        }

        /// <summary>
        /// Splits a string into parts, grouping letters, numbers and punctuation
        /// </summary>
        /// <param name="StringToSplit">The string to split.</param>
        /// <returns></returns>
        public static string[] SplitString(string StringToSplit)
        {
            //TODO - Move this out of Excel functions to SNT Tools or other place
            var mySplit = new List<string>();
            List<char> currString = new List<char>();
            char currChar;
            CharType lastCharType = CharType.Unknown;
            CharType thisCharType = CharType.Unknown;

            char[] myChars = StringToSplit.ToCharArray();
            for (int i = 0; i < myChars.Length; i++)
            {
                currChar = myChars[i];
                switch (lastCharType)
                {
                    case CharType.Unknown:
                        currString.Add(currChar);

                        if (char.IsNumber(currChar))
                            thisCharType = CharType.Number;
                        else if (char.IsLetter(currChar))
                            thisCharType = CharType.Letter;
                        else
                            thisCharType = CharType.Punctuation;

                        break;

                    case CharType.Number:
                        if (char.IsNumber(currChar))
                            thisCharType = CharType.Number;
                        else if (char.IsLetter(currChar))
                            thisCharType = CharType.Letter;
                        else
                            thisCharType = CharType.Punctuation;

                        switch (thisCharType)
                        {
                            case CharType.Number:
                                currString.Add(currChar);
                                break;

                            case CharType.Letter:
                                mySplit.Add(new string(currString.ToArray()));

                                currString.Clear();
                                currString.Add(currChar);
                                break;

                            case CharType.Punctuation:
                                mySplit.Add(new string(currString.ToArray()));

                                currString.Clear();
                                currString.Add(currChar);
                                break;

                            default:
                                break;
                        }
                        break;

                    case CharType.Letter:
                        if (char.IsNumber(currChar))
                            thisCharType = CharType.Number;
                        else if (char.IsLetter(currChar))
                            thisCharType = CharType.Letter;
                        else
                            thisCharType = CharType.Punctuation;

                        switch (thisCharType)
                        {
                            case CharType.Number:
                                mySplit.Add(new string(currString.ToArray()));

                                currString.Clear();
                                currString.Add(currChar);
                                break;

                            case CharType.Letter:
                                currString.Add(currChar);
                                break;

                            case CharType.Punctuation:
                                mySplit.Add(new string(currString.ToArray()));

                                currString.Clear();
                                currString.Add(currChar);
                                break;

                            default:
                                break;
                        }
                        break;

                    case CharType.Punctuation:
                        if (char.IsNumber(currChar))
                            thisCharType = CharType.Number;
                        else if (char.IsLetter(currChar))
                            thisCharType = CharType.Letter;
                        else
                            thisCharType = CharType.Punctuation;

                        switch (thisCharType)
                        {
                            case CharType.Number:
                                mySplit.Add(new string(currString.ToArray()));

                                currString.Clear();
                                currString.Add(currChar);
                                break;

                            case CharType.Letter:
                                mySplit.Add(new string(currString.ToArray()));

                                currString.Clear();
                                currString.Add(currChar);
                                break;

                            case CharType.Punctuation:
                                currString.Add(currChar);
                                break;

                            default:
                                break;
                        }
                        break;

                    default:
                        break;
                }

                lastCharType = thisCharType;
            }

            mySplit.Add(new string(currString.ToArray()));
            return mySplit.ToArray();
        }

        #region Methods

        /// <summary>
        /// Converts an Excel Cell value to Double. Excel cells are typically stored as string, double or null.
        /// </summary>
        /// <param name="Cell"></param>
        /// <returns>The double value or 0 if null.</returns>
        public static double ExcelCellToDouble(Range Cell)
        {
            try
            {
                if (Cell.Value is null) return 0.0;
                if (Cell.Value is double) return Convert.ToDouble(Cell.Value);
                if (Cell.Value is string && string.IsNullOrWhiteSpace(Cell.Value)) return 0.0;
                return Convert.ToDouble(Cell.Value); // Try conversion or fail
            }
            catch (Exception ex)
            {
                Trace.TraceError($"Cannot convert valve '{Cell.Value}' in cell {Cell.Address} in sheet {Cell.Worksheet.Name} to Double: {ex.Message}");
                return 0.0;
            }
        }

        /// <summary>
        /// Converts an Excel Cell value to Int32. Excel cells are typically stored as string, double or null.
        /// </summary>
        /// <param name="Cell"></param>
        /// <returns>The integer value or 0 if null.</returns>
        public static int ExcelCellToInt32(Range Cell)
        {
            try
            {
                if (Cell.Value is null) return 0;
                if (Cell.Value is double) return Convert.ToInt32(Cell.Value);
                if (Cell.Value is string && string.IsNullOrWhiteSpace(Cell.Value)) return 0;
                return Convert.ToInt32(Cell.Value); // Try conversion or fail
            }
            catch (Exception ex)
            {
                Trace.TraceError($"Cannot convert valve '{Cell.Value}' in cell {Cell.Address} in sheet {Cell.Worksheet.Name} to Int32: {ex.Message}");
                return 0;
            }
        }

        /// <summary>
        /// Converts an Excel Cell value to string. Excel cells are typically stored as string, double or null.
        /// </summary>
        /// <param name="Cell"></param>
        /// <returns>The string value.</returns>
        public static string ExcelCellToString(Range Cell)
        {
            try
            {
                return Cell.Value?.ToString() ?? "";
            }
            catch (Exception ex)
            {
                Trace.TraceError($"Cannot convert valve '{Cell.Value}' in cell {Cell.Address} in sheet {Cell.Worksheet.Name} to String: {ex.Message}");
                return "";
            }
        }

        public static int FindColumnHeader(Range Row, string ColumnName)
        {
            string colName = "";
            for (int i = 1; i <= Row.Columns.Count; i++)
            {
                colName = Row.Columns[i].Value?.ToString() ?? "";
                if (colName == "")
                    colName = $"Column{ColumnLetter(i, false)}";

                if (ColumnName.Equals(colName, StringComparison.OrdinalIgnoreCase))
                    return i;
            }

            // Column header was not found
            Trace.TraceWarning($"Column header required '{ColumnName}' does not exist in the table. A list of column headers follows...");
            for (int i = 1; i <= Row.Columns.Count; i++)
            {
                Trace.TraceWarning($"Detected column header: '{Row.Columns[i].Value?.ToString()}'.");
            }
            return -1;
        }

        #endregion Methods

        /// <summary>
        /// Returns the column letter from the corresponding column index.
        /// </summary>
        /// <param name="ColumnIndex">The column index.</param>
        /// <param name="StartAtZero">If the column index is 0 or 1 based.</param>
        /// <returns></returns>
        public static string ColumnLetter(int ColumnIndex, bool StartAtZero)
        {
            if (ColumnIndex == 0 && !StartAtZero)
                return "";

            int colIndex = StartAtZero ? ColumnIndex : ColumnIndex - 1;

            string[] columns = new string[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J",
            "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z",
            "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM",
            "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ",
            "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM",
            "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ", };

            if (colIndex >= columns.Length)
                return "";

            return columns[colIndex];
        }
    }

    public enum ExcelCol
    {
        A = 0,
        B = 1,
        C = 2,
        D = 3,
        E = 4,
        F = 5,
        G = 6,
        H = 7,
        I = 8,
        J = 9,
        K = 10,
        L = 11,
        M = 12,
        N = 13,
        O = 14,
        P = 15,
        Q = 16,
        R = 17,
        S = 18,
        T = 19,
        U = 20,
        V = 21,
        W = 22,
        X = 23,
        Y = 24,
        Z = 25,
        AA = 26,
        AB = 27,
        AC = 28,
        AD = 29,
        AE = 30,
        AF = 31,
        AG = 32,
        AH = 33,
        AI = 34,
        AJ = 35,
        AK = 36,
        AL = 37,
        AM = 38,
        AN = 39,
        AO = 40,
        AP = 41,
        AQ = 42,
        AR = 43,
        AS = 44,
        AT = 45,
        AU = 46,
        AV = 47,
        AW = 48,
        AX = 49,
        AY = 50,
        AZ = 51,
        BA = 52,
        BB = 53,
        BC = 54,
        BD = 55,
        BE = 56,
        BF = 57,
        BG = 58,
        BH = 59,
        BI = 60,
        BJ = 61,
        BK = 62,
        BL = 63,
        BM = 64,
        BN = 65,
        BO = 66,
        BP = 67,
        BQ = 68,
        BR = 69,
        BS = 70,
        BT = 71,
        BU = 72,
        BV = 73,
        BW = 74,
        BX = 75,
        BY = 76,
        BZ = 77
    }
}