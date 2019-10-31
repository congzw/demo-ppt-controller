using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace OfficeCtrlDemo
{
    public static class IniStringExtensions
    {
        public static IDictionary<string, object> AsIniDic(this string iniString, string equalOperator = "=", char lastSplit = ';')
        {
            var result = new Dictionary<string, object>();

            if (string.IsNullOrWhiteSpace(iniString))
            {
                return result;
            }

            var values = iniString.Split(lastSplit);
            foreach (var value in values)
            {
                var items = value.Split(new string[] { equalOperator }, StringSplitOptions.None);
                var propName = items[0];
                object propValue = null;
                if (items.Length == 2)
                {
                    if (!string.IsNullOrWhiteSpace(items[1]))
                    {
                        propValue = items[1].Trim();
                    }
                }
                result.Add(propName, propValue);
            }
            return result;
        }

        public static string AsIniString(this object instance, IEnumerable<string> ignoredNames, string equalOperator = "=", char lastSplit = ';', bool removeLastSplit = true)
        {
            if (instance == null)
            {
                return string.Empty;
            }

            var dic = new Dictionary<string, object>();
            var propertyInfos = instance.GetType().GetProperties(BindingFlags.Instance | BindingFlags.Public);
            foreach (PropertyInfo propInfo in propertyInfos)
            {
                object value = propInfo.GetValue(instance, null);
                dic.Add(propInfo.Name, value);
            }
            return dic.AsIniString(ignoredNames, equalOperator, lastSplit, removeLastSplit);
        }

        public static string AsIniString(this IDictionary<string, object> dic, IEnumerable<string> ignoredNames, string equalOperator = "=", char lastSplit = ';', bool removeLastSplit = true)
        {
            if (dic == null || dic.Count == 0)
            {
                return string.Empty;
            }

            var fixIgnoredNames = new List<string>();
            if (ignoredNames != null)
            {
                fixIgnoredNames.AddRange(ignoredNames);
            }

            var schema = string.Format("{0}{1}{2}{3}", "{0}", equalOperator, "{1}", lastSplit);
            var sb = new StringBuilder();
            foreach (var item in dic)
            {
                if (fixIgnoredNames.Any(x => x.Equals(item.Key, StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }
                var temp = item.Value == null ? string.Empty : item.Value.ToString();
                sb.AppendFormat(schema, item.Key, temp);
            }

            //去掉最后的分号
            string result = sb.ToString();
            if (removeLastSplit && !string.IsNullOrWhiteSpace(result))
            {
                return result.Substring(0, result.Length - 1);
            }

            return sb.ToString();
        }
    }
}
