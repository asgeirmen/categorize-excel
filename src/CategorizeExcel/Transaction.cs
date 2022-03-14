using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CategorizeExcel
{
    public class Transaction
    {
        public string Id { get; set; }

        public DateTime TransactionDateTime { get; set; }

        public string Text { get; set; }

        public int CategoryId { get; set; }
    }
}
