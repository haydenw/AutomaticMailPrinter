using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomaticMailPrinter
{
    public class Order
    {
        public int id;
        public string subject;
        public DateTime createdAt;
        public DateTime? printedAt;
        public string html;
    }
}
