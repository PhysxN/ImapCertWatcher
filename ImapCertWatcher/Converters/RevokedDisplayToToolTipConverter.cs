using System;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Windows.Data;

namespace ImapCertWatcher.Converters
{
    public sealed class RevokedDisplayToToolTipConverter : IValueConverter
    {
        private static readonly Regex DateRegex =
            new Regex(@"\b\d{2}\.\d{2}\.\d{4}\b", RegexOptions.Compiled);

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var text = value as string;

            if (string.IsNullOrWhiteSpace(text))
                return "Аннулирован";

            var match = DateRegex.Match(text);
            if (match.Success)
                return "Аннулирован: " + match.Value;

            return "Аннулирован";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }
    }
}