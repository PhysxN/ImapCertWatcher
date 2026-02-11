using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Data;
using ImapCertWatcher.Models;

namespace ImapCertWatcher
{
    public class TokensForCertConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            var cert = values[0] as CertRecord;
            var allTokens = values[1] as IEnumerable<TokenRecord>;

            if (allTokens == null)
                return null;

            if (cert == null)
                return allTokens.Where(t => t.IsFree);

            // свободные + уже привязанный к сертификату
            return allTokens.Where(t => t.IsFree || t.Id == cert.TokenId).ToList();
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
            => throw new NotSupportedException();
    }
}
