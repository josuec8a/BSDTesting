using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Threading.Tasks;

namespace BSD_DataProcessing.Helpers
{
    public class MyWebClient : WebClient
    {
        public MyWebClient(int timeout)
        {
            this.Timeout = timeout;
        }
        public int Timeout { get; set; }
        protected override WebRequest GetWebRequest(Uri address)
        {
            WebRequest _webRequest = base.GetWebRequest(address);
            _webRequest.Timeout = Timeout;
            ((HttpWebRequest)_webRequest).ReadWriteTimeout = Timeout;

            return _webRequest;
        }
    }
}
