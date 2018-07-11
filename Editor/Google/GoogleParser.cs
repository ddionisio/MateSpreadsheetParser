using UnityEngine;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.IO;
using System;
using System.Linq;
using System.ComponentModel;

using Google.GData.Client;
using Google.GData.Spreadsheets;

namespace M8.SpreadsheetParser {
    public class GoogleParser {
        private string mSheetName;
        //private DatabaseClient mClient;

        private GoogleDatabaseQuery mDatabaseQuery;
        private GoogleDatabase mDatabase;

        public GoogleParser(GoogleSettings settings, string sheetName) {
            mSheetName = sheetName;
            mDatabaseQuery = new GoogleDatabaseQuery(settings);
        }

        public bool GenerateDatabase(ref string error) {
            mDatabase = mDatabaseQuery.GetDatabase(mSheetName, ref error);
            if(!string.IsNullOrEmpty(error))
                return false;

            if(mDatabase.worksheets.Length == 0) {
                error = "No Worksheets found in: " + mSheetName;
                return false;
            }

            return string.IsNullOrEmpty(error);
        }

        /// <summary>
        /// Deserialize all sheets
        /// </summary>
        public Dictionary<string, List<T>> DeserializeAllSheets<T>(string[] filterFieldNames) {
            var convertedSheets = new Dictionary<string, List<T>>();

            for(int i = 0; i < mDatabase.worksheets.Length; i++) {
                var sheet = mDatabase.worksheets[i];
                var sheetName = sheet.Title.Text;

                var convertList = Deserialize<T>(sheet, filterFieldNames);

                convertedSheets.Add(sheetName, convertList);
            }

            return convertedSheets;
        }

        //public List<T> 
        /// <summary>
        /// Deserialize all rows starting at startRow. First row defines which fields to apply, optional filterFieldNames to only grab from those columns
        /// </summary>
        public List<T> Deserialize<T>(int sheetIndex, string[] filterFieldNames) {
            var sheet = mDatabase.worksheets[sheetIndex];
            return Deserialize<T>(sheet, filterFieldNames);
        }

        protected List<T> Deserialize<T>(WorksheetEntry sheet, string[] filterFieldNames) {
            var output = new List<T>();

            HashSet<string> filterFieldLookup = filterFieldNames != null && filterFieldNames.Length > 0 ? new HashSet<string>(filterFieldNames) : null;

            //grab headers
            var headerQuery = new CellQuery(sheet.CellFeedLink);
            headerQuery.MinimumRow = 1;
            headerQuery.MaximumRow = 1;

            var headerCellFeed = mDatabaseQuery.spreadsheetService.Query(headerQuery) as CellFeed;
            //

            //grab contents
            var contentQuery = new CellQuery(sheet.CellFeedLink);
            contentQuery.MinimumRow = 2;

            var contentFeed = mDatabaseQuery.spreadsheetService.Query(contentQuery) as CellFeed;

            uint curRow = 0;

            var type = typeof(T);

            T item = default(T);

            foreach(CellEntry cell in contentFeed.Entries) {
                //next item?
                if(curRow != cell.Row) {
                    item = Activator.CreateInstance<T>();
                    output.Add(item);

                    curRow = cell.Row;
                }

                if(string.IsNullOrEmpty(cell.Value))
                    continue;

                //get associated header
                var headerCell = headerCellFeed[1, cell.Column];
                if(headerCell == null)
                    continue;

                string headerName = headerCell.Value;
                if(string.IsNullOrEmpty(headerName))
                    continue;

                if(filterFieldLookup != null && !filterFieldLookup.Contains(headerName))
                    continue;

                //set value

                const BindingFlags bindFlags = BindingFlags.IgnoreCase | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance;

                //try field
                FieldInfo fieldInfo = type.GetField(headerName, bindFlags);
                if(fieldInfo != null) {
                    try {
                        var value = ConvertFrom(cell, fieldInfo.FieldType);
                        fieldInfo.SetValue(item, value);
                    }
                    catch(Exception e) {
                        string pos = string.Format("Row[{0}] Col[{1}]", cell.Row, cell.Column);
                        Debug.LogError(string.Format("Deserialize {0}:{1} ({2})  Exception: {3}", mSheetName, sheet.Title.Text, pos, e.Message));
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
                        string pos = string.Format("Row[{0}] Col[{1}]", cell.Row, cell.Column);
                        Debug.LogError(string.Format("Deserialize {0}:{1} ({2})  Exception: {3}", mSheetName, sheet.Title.Text, pos, e.Message));
                    }

                    continue;
                }
            }
            //

            return output;
        }

        protected object ConvertFrom(CellEntry cell, Type t) {
            object value = null;

            if(t == typeof(float) || t == typeof(double) || t == typeof(short) || t == typeof(int) || t == typeof(long)) {
                string cellVal = !string.IsNullOrEmpty(cell.NumericValue) ? cell.NumericValue : cell.Value;

                if(t == typeof(float))
                    value = Convert.ToSingle(cellVal);
                if(t == typeof(double))
                    value = Convert.ToDouble(cellVal);
                if(t == typeof(short))
                    value = Convert.ToInt16(cellVal);
                if(t == typeof(int))
                    value = Convert.ToInt32(cellVal);
                if(t == typeof(long))
                    value = Convert.ToInt64(cellVal);
            }
            else if(t == typeof(string) || t.IsArray) {
                value = cell.Value;
            }
            else if(t == typeof(bool)) {
                string cellVal = cell.Value.ToLower();
                value = cellVal == "true" || cellVal == "1";
            }

            if(t.IsGenericType && t.GetGenericTypeDefinition().Equals(typeof(Nullable<>))) {
                var nc = new NullableConverter(t);
                return nc.ConvertFrom(value);
            }

            if(t.IsEnum) {
                // for enum type, first get value by string then convert it to enum.
                return Enum.Parse(t, cell.Value.Trim(), true);
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