using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImapCertWatcher.Models
{
    public class TokenRecord
    {
        public int Id { get; set; }
        public string Sn { get; set; }
        public int? OwnerCertId { get; set; }

        // для UI
        public string OwnerFio { get; set; }

        public bool IsFree => OwnerCertId == null;

        // 🔽 ДОБАВЬ ЭТО СВОЙСТВО
        public string OwnerFioFormatted
        {
            get
            {
                if (string.IsNullOrWhiteSpace(OwnerFio))
                    return OwnerFio;

                var words = OwnerFio
                    .ToLower()
                    .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                for (int i = 0; i < words.Length; i++)
                {
                    if (words[i].Length > 1)
                        words[i] = char.ToUpper(words[i][0]) + words[i].Substring(1);
                    else
                        words[i] = words[i].ToUpper();
                }

                return string.Join(" ", words);
            }
        }
    }
}