using System;
using System.Globalization;
using System.Windows.Data;

namespace ImapCertWatcher.Utils
{
    public class NameCaseConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null)
                return "";

            string text = value.ToString().Trim();

            if (string.IsNullOrEmpty(text))
                return text;

            // "ИВАНОВ ИВАН ИВАНОВИЧ" -> "Иванов Иван Иванович"
            text = text.ToLower();

            var words = text.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            for (int i = 0; i < words.Length; i++)
            {
                if (words[i].Length > 1)
                {
                    words[i] =
                        char.ToUpper(words[i][0]) +
                        words[i].Substring(1);
                }
                else
                {
                    words[i] = words[i].ToUpper();
                }
            }

            return string.Join(" ", words);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value;
        }
    }
}