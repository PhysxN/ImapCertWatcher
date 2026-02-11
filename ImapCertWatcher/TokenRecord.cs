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
    }
}