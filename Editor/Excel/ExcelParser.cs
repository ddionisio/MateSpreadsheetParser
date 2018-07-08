using UnityEngine;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.IO;
using System;
using System.Linq;
using System.ComponentModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace M8.SpreadsheetParser {
    public class ExcelParser {
        public int sheetCount { get { return mWorkbook != null ? mWorkbook.NumberOfSheets : 0; } }

        private IWorkbook mWorkbook;

        private string mFilepath;

        public ExcelParser(string filepath) {
            try {
                using(FileStream fileStream = new FileStream(filepath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) {
                    mFilepath = filepath;

                    string extension = mFilepath.GetExtension();

                    if(extension == "xls")
                        mWorkbook = new HSSFWorkbook(fileStream);
                    else if(extension == "xlsx") {
#if UNITY_EDITOR_OSX
                        throw new Exception("xlsx is not supported on OSX.");
#else
                        mWorkbook = new XSSFWorkbook(fileStream);
#endif
                    }
                    else {
                        throw new Exception("Invalid file: " + mFilepath);
                    }
                }
            }
            catch(Exception e) {
                Debug.LogError(e.Message);
            }
        }

        public string GetSheetName(int sheetIndex) {
            return mWorkbook.GetSheetName(sheetIndex);
        }

        public string[] GetSheetNames() {
            var sheetNames = new string[sheetCount];
            for(int i = 0; i < sheetNames.Length; i++)
                sheetNames[i] = mWorkbook.GetSheetName(i);
            return sheetNames;
        }

        /// <summary>
        /// Grab from first sheet, set row=0 to grab "headers". Empty cells are not added.
        /// </summary>
        public string[] GetRowStrings(int row) {
            return GetRowStrings(0, row);
        }

        /// <summary>
        /// Grab from given sheet index, set row=0 to grab "headers". Empty cells are not added.
        /// </summary>
        public string[] GetRowStrings(int sheetIndex, int row) {
            List<string> result = new List<string>();

            var sheet = mWorkbook.GetSheetAt(sheetIndex);
            var sheetRow = sheet.GetRow(row);
            if(sheetRow != null) {
                for(int i = 0; i < sheetRow.LastCellNum; i++) {
                    string value = sheetRow.GetCell(i).StringCellValue;
                    if(!string.IsNullOrEmpty(value))
                        result.Add(value);
                }
            }

            return result.ToArray();
        }

        /// <summary>
        /// Deserialize all sheets
        /// </summary>
        public Dictionary<string, List<T>> Deserializer<T>(string[] filterFieldNames) {
            var convertedSheets = new Dictionary<string, List<T>>();

            for(int i = 0; i < sheetCount; i++) {
                var sheetName = mWorkbook.GetSheetName(i);
                var sheet = mWorkbook.GetSheetAt(i);

                var convertList = Deserialize<T>(sheet, filterFieldNames);

                convertedSheets.Add(sheetName, convertList);
            }

            return convertedSheets;
        }

        /// <summary>
        /// Deserialize all rows starting at startRow. First row defines which fields to apply, optional filterFieldNames to only grab from those columns
        /// </summary>
        public List<T> Deserialize<T>(int sheetIndex, string[] filterFieldNames) {
            var sheet = mWorkbook.GetSheetAt(sheetIndex);
            return Deserialize<T>(sheet, filterFieldNames);
        }


        protected List<T> Deserialize<T>(ISheet sheet, string[] filterFieldNames) {
            var type = typeof(T);

            var result = new List<T>();

            List<string> headerNames = new List<string>();

            HashSet<string> filterFieldLookup = filterFieldNames != null && filterFieldNames.Length > 0 ? new HashSet<string>(filterFieldNames) : null;

            int rowIndex = 0;
            foreach(IRow row in sheet) {
                if(rowIndex == 0) {
                    //setup header names
                    for(int i = 0; i < row.LastCellNum; i++) {
                        string value = row.GetCell(i).StringCellValue;
                        headerNames.Add(value);
                    }
                }
                else {
                    var item = (T)Activator.CreateInstance(type);

                    //fill item with values from row
                    //grab cells based on headerNames, filter with fieldLookup (if valid)
                    for(int headerIndex = 0; headerIndex < headerNames.Count; headerIndex++) {
                        string headerName = headerNames[headerIndex];

                        if(filterFieldLookup != null && !filterFieldLookup.Contains(headerName))
                            continue;

                        var cell = row.GetCell(headerIndex, MissingCellPolicy.RETURN_BLANK_AS_NULL);
                        if(cell == null)
                            continue;

                        const BindingFlags bindFlags = BindingFlags.IgnoreCase | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance;

                        //try field
                        FieldInfo fieldInfo = type.GetField(headerName, bindFlags);
                        if(fieldInfo != null) {
                            try {
                                var value = ConvertFrom(cell, fieldInfo.FieldType);
                                fieldInfo.SetValue(item, value);
                            }
                            catch(Exception e) {
                                string pos = string.Format("Row[{0}], Cell[{1}]", (rowIndex).ToString(), headerNames[headerIndex]);
                                Debug.LogError(string.Format("Deserialize {0} ({1})  Exception: {2}", mFilepath, pos, e.Message));
                            }

                            continue;
                        }

                        //try property
                        PropertyInfo propInfo = type.GetProperty(headerName, bindFlags);
                        if(propInfo != null) {
                            if(!propInfo.CanWrite) //ignore readonly properties
                                continue;

                            try {
                                var value = ConvertFrom(cell, propInfo.PropertyType);
                                propInfo.SetValue(item, value, null);
                            }
                            catch(Exception e) {
                                string pos = string.Format("Row[{0}], Cell[{1}]", (rowIndex).ToString(), headerNames[headerIndex]);
                                Debug.LogError(string.Format("Deserialize {0} ({1})  Exception: {2}", mFilepath, pos, e.Message));
                            }

                            continue;
                        }
                    }

                    result.Add(item);
                }

                rowIndex++;
            }

            return result;
        }

        protected object ConvertFrom(ICell cell, Type t) {
            object value = null;

            if(t == typeof(float) || t == typeof(double) || t == typeof(short) || t == typeof(int) || t == typeof(long)) {
                if(cell.CellType == NPOI.SS.UserModel.CellType.Numeric) {
                    value = cell.NumericCellValue;
                }
                else if(cell.CellType == NPOI.SS.UserModel.CellType.String) {
                    //Get correct numeric value even the cell is string type but defined with a numeric type in a data class.
                    if(t == typeof(float))
                        value = Convert.ToSingle(cell.StringCellValue);
                    if(t == typeof(double))
                        value = Convert.ToDouble(cell.StringCellValue);
                    if(t == typeof(short))
                        value = Convert.ToInt16(cell.StringCellValue);
                    if(t == typeof(int))
                        value = Convert.ToInt32(cell.StringCellValue);
                    if(t == typeof(long))
                        value = Convert.ToInt64(cell.StringCellValue);
                }
                else if(cell.CellType == NPOI.SS.UserModel.CellType.Formula) {
                    // Get value even if cell is a formula
                    if(t == typeof(float))
                        value = Convert.ToSingle(cell.NumericCellValue);
                    if(t == typeof(double))
                        value = Convert.ToDouble(cell.NumericCellValue);
                    if(t == typeof(short))
                        value = Convert.ToInt16(cell.NumericCellValue);
                    if(t == typeof(int))
                        value = Convert.ToInt32(cell.NumericCellValue);
                    if(t == typeof(long))
                        value = Convert.ToInt64(cell.NumericCellValue);
                }
            }
            else if(t == typeof(string) || t.IsArray) {
                // HACK: handles the case that a cell contains numeric value
                //       but a member field in a data class is defined as string type.
                //       e.g. string s = "123"
                if(cell.CellType == NPOI.SS.UserModel.CellType.Numeric)
                    value = cell.NumericCellValue;
                else
                    value = cell.StringCellValue;
            }
            else if(t == typeof(bool))
                value = cell.BooleanCellValue;

            if(t.IsGenericType && t.GetGenericTypeDefinition().Equals(typeof(Nullable<>))) {
                var nc = new NullableConverter(t);
                return nc.ConvertFrom(value);
            }

            if(t.IsEnum) {
                // for enum type, first get value by string then convert it to enum.
                value = cell.StringCellValue;
                return Enum.Parse(t, value.ToString(), true);
            }
            else if(t.IsArray) {
                var valueStr = (string)value;

                if(t.GetElementType() == typeof(float))
                    return valueStr.ToFloatArray();

                if(t.GetElementType() == typeof(double))
                    return valueStr.ToDoubleArray();

                if(t.GetElementType() == typeof(short))
                    return valueStr.ToShortArray();

                if(t.GetElementType() == typeof(int))
                    return valueStr.ToIntArray();

                if(t.GetElementType() == typeof(long))
                    return valueStr.ToLongArray();

                if(t.GetElementType() == typeof(string))
                    return valueStr.ToStringArray();
            }

            // for all other types, convert its corresponding type.
            return Convert.ChangeType(value, t);
        }
    }
}