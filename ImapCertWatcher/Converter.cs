using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Data;
using ImapCertWatcher.Models;

public class TokensForCertConverter : IMultiValueConverter
{
    public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
    {
        if (values.Length != 2)
            return null;

        var cert = values[0] as CertRecord;
        var tokens = values[1] as IEnumerable<TokenRecord>;

        if (tokens == null)
            return null;

        if (cert == null || cert.TokenId == null)
            return tokens.Where(t => t.IsFree).ToList();

        return tokens.Where(t =>
            t.IsFree || t.Id == cert.TokenId).ToList();
    }

    public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        => throw new NotSupportedException();
}
