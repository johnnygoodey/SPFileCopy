using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Britehouse.SPFileCopy
{
    public class Mapping
    {
        public string Site { get; set; }
        public string Source { get; set; }
        public string Target { get; set; }

        public string FilePath { get; set; }

        public string ProjectFullPath { get; set; }
        
        public string FullTargetFolder
        {
            get
            {
                var sourceLocation = string.Format("{0}\\{1}", Path.GetDirectoryName(ProjectFullPath), Source);
                sourceLocation = sourceLocation.TrimEnd('\\');
                var targetLocation = Target.TrimEnd('/');

                return ReplaceCaseInsensitive(FilePath, sourceLocation, targetLocation);
            }
        }

        public bool Checkout { get; set; }
        public bool Publish { get; set; }
        public bool Approve { get; set; }

        private static string ReplaceCaseInsensitive(string input, string search, string replacement)
        {
            var inputString = input.Replace("\\", "/");
            var replace = search.Replace("\\", "/");
            string result = Regex.Replace(
                inputString,
                Regex.Escape(replace),
                replacement.Replace("$", "$$"),
                RegexOptions.IgnoreCase
            );
            return result;
        }
    }
}
