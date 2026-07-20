
using System;
using System.Linq;

namespace Contensive.Addons.Newsletter.Controllers {
    public class GenericController {
        // 
        // ====================================================================================================
        /// <summary>
        /// return date.minValue if date is before 1/1/1900
        /// </summary>
        /// <param name="sourceDate"></param>
        /// <returns></returns>
        public static DateTime encodeMinDate(DateTime sourceDate) {
            if (sourceDate < DateTime.Parse("1900-01-01")) {
                return DateTime.MinValue;
            }
            return sourceDate;
        }

        public static bool isNumeric(string value) {
            return value.All(char.IsNumber);
        }

        public static bool isDateEmpty(DateTime srcDate) {
            return srcDate < new DateTime(1900, 1, 1);
        }

        public static string getShortDateString(DateTime srcDate) {
            if (!isDateEmpty(srcDate)) {
                return encodeMinDate(srcDate).ToShortDateString();
            }

            return string.Empty;
        }

        public static string getSortOrderFromInteger(int id) {
            return id.ToString().PadLeft(7, '0');
        }

        public static string getDateForHtmlInput(DateTime source) {
            if (isDateEmpty(source)) {
                return "";
            } else {
                return source.Year + "-" + source.Month.ToString().PadLeft(2, '0') + "-" + source.Day.ToString().PadLeft(2, '0');
            }
        }

        public static string verifyProtocol(string url) {
            if (string.IsNullOrWhiteSpace(url))
                return string.Empty;
            if (url.Substring(0, 1) == "/")
                return url;
            if (!url.IndexOf("://").Equals(-1))
                return url;
            return "http://" + url;
        }

    }
}