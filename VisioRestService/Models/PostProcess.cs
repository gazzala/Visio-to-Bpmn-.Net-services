using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace VisioRestService.Models
{
    public class PostProcess
    {

        public string processName { get; set; }

        public List<ExportElement> elementDetails;
    }
}