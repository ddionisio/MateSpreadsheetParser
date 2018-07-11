using System;
using System.Collections;
using System.Collections.Generic;

namespace M8.SpreadsheetParser {
    public static class StringExt {
        private static readonly char[] delimiters = new char[] { ',', ';' };

        /// <summary>
        /// Grab extension if valid (filename.ext = ext), return empty otherwise.
        /// </summary>
        public static string GetExtension(this string str) {
            int extInd = str.LastIndexOf('.');
            if(extInd == -1 || extInd == str.Length - 1)
                return "";

            return str.Substring(extInd + 1);
        }

        public static float[] ToFloatArray(this string str) {
            var parseItems = str.ToStringArray();

            var valueList = new List<float>(parseItems.Length);

            for(int i = 0; i < parseItems.Length; i++) {
                float val;
                if(float.TryParse(parseItems[i], out val))
                    valueList.Add(val);
                else
                    valueList.Add(0f);
            }

            return valueList.ToArray();
        }

        public static double[] ToDoubleArray(this string str) {
            var parseItems = str.ToStringArray();

            var valueList = new List<double>(parseItems.Length);

            for(int i = 0; i < parseItems.Length; i++) {
                double val;
                if(double.TryParse(parseItems[i], out val))
                    valueList.Add(val);
                else
                    valueList.Add(0.0);
            }

            return valueList.ToArray();
        }

        public static int[] ToIntArray(this string str) {
            var parseItems = str.ToStringArray();

            var valueList = new List<int>(parseItems.Length);

            for(int i = 0; i < parseItems.Length; i++) {
                int val;
                if(int.TryParse(parseItems[i], out val))
                    valueList.Add(val);
                else
                    valueList.Add(0);
            }

            return valueList.ToArray();
        }

        public static short[] ToShortArray(this string str) {
            var parseItems = str.ToStringArray();

            var valueList = new List<short>(parseItems.Length);

            for(int i = 0; i < parseItems.Length; i++) {
                short val;
                if(short.TryParse(parseItems[i], out val))
                    valueList.Add(val);
                else
                    valueList.Add(0);
            }

            return valueList.ToArray();
        }

        public static long[] ToLongArray(this string str) {
            var parseItems = str.ToStringArray();

            var valueList = new List<long>(parseItems.Length);

            for(int i = 0; i < parseItems.Length; i++) {
                long val;
                if(long.TryParse(parseItems[i], out val))
                    valueList.Add(val);
                else
                    valueList.Add(0);
            }

            return valueList.ToArray();
        }

        public static string[] ToStringArray(this string str) {
            var splits = str.Split(delimiters, StringSplitOptions.RemoveEmptyEntries);

            for(int i = 0; i < splits.Length; i++)
                splits[i] = splits[i].Trim();

            return splits;
        }
    }
}