using System;
using System.Globalization;
using System.Windows.Data;

namespace ImapCertWatcher.Utils

{
    public class BoolToFreeBusyConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return (value is bool b && b) ? "Свободен" : "Занят";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
