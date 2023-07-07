using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace JDETakeMail
{
    static class Extensions
    {
        public static string Right(this string value, int length)
        {
            if (value.Length < length)
                return value;
            return value.Substring(value.Length - length);
        }

        public static bool IsNumber(this string value)
        {
            //return !Regex.IsMatch(value, "[^0-9]");
            return int.TryParse(value, out _);
        }
    }
}
