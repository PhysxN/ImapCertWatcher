using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImapCertWatcher.Models
{
    public class TokenListItem
    {
        public bool IsHeader { get; set; }
        public string HeaderText { get; set; }
        public TokenRecord Token { get; set; }

        public string GroupName => Token.IsFree ? "Свободные" : "Занятые";

    }
}
