using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MsDyn.Contrib.CloneFieldDefinitions
{
    public static class StringExtensions
    {
        public static string ReplaceEntityName(this string name, string source, string target)
        {
            var replacedName = string.Empty;

            if (string.IsNullOrEmpty(name))
            {
                return replacedName;
            }

            var sourcePattern = $"_{source}_";
            if (name.Contains(sourcePattern))
            {
                replacedName = name.Replace(sourcePattern, $"_{target}_");
            }
            else
            {
                replacedName = $"{name}_{target}";
            }

            return replacedName;
        }
    }
}
