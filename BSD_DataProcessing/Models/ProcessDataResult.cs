using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace BSD_DataProcessing.Models
{
    public class ProcessDataResult
    {
        public DataTable DataProcessed { get; set; }
        public List<string> RetryList { get; set; }
    }
}
