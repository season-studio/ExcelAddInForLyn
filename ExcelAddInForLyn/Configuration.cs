using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddInForLyn
{
    internal static class Configuration
    {
        internal static readonly string ConfigurationKey = @"software\season-studio\ExcalAddinForLyn";

        public static void Set(string _name, object _val)
        {
            var key = Registry.CurrentUser.CreateSubKey(ConfigurationKey, true);
            key.SetValue(_name, _val);
        }

        public static object Get(string _name, object _default)
        {
            var key = Registry.CurrentUser.CreateSubKey(ConfigurationKey, false);
            return key.GetValue(_name, _default);
        }
    }
}
