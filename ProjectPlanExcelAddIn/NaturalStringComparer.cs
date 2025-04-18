using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

public class NaturalStringComparer : IComparer<string>, IComparer
{
    public int Compare(string x, string y)
    {
        return CompareNatural(x, y);
    }

    public int Compare(object x, object y)
    {
        return Compare(x?.ToString(), y?.ToString());
    }

    private static int CompareNatural(string a, string b)
    {
        if (a == b)
            return 0;
        if (a == null)
            return -1;
        if (b == null)
            return 1;

        var regex = new Regex(@"\d+|\D+");
        var aParts = regex.Matches(a);
        var bParts = regex.Matches(b);
        int count = Math.Min(aParts.Count, bParts.Count);

        for (int i = 0; i < count; i++)
        {
            string aPart = aParts[i].Value;
            string bPart = bParts[i].Value;

            if (int.TryParse(aPart, out int aNum) && int.TryParse(bPart, out int bNum))
            {
                int numCompare = aNum.CompareTo(bNum);
                if (numCompare != 0)
                    return numCompare;
            }
            else
            {
                int strCompare = string.Compare(aPart, bPart, CultureInfo.CurrentCulture, CompareOptions.IgnoreCase);
                if (strCompare != 0)
                    return strCompare;
            }
        }

        return aParts.Count.CompareTo(bParts.Count);
    }
}
