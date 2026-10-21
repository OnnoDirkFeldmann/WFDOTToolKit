using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WFDOTToolKit
{
    internal class PluralShorter
    {
        internal static string GetShortPlural(string word, string plural)
        {
            int substring = 0;
            for (int i = 0; i < plural.Length; i++)
            {
                if (i >= word.Length || !plural[i].Equals(word[i]))
                {
                    substring = i;
                    break;
                }
            }
            if (substring < 3) return plural;
            var shortplural = plural.Substring(substring, plural.Length - substring);
            return $"-{shortplural}";
        }
    }
}
