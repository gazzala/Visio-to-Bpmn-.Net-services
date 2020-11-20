using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Web;

namespace VisioRestService.Models
{
    public class ExportElement
    {

        public string ShapeID { get; set; }

        public string elementName { get; set; }

        public string elementType { get; set; }

        public string fromID { get; set; }
        
        public string condition { get; set; } = string.Empty;

        [JsonIgnore]
        public long SeqID { get; set; }
        [JsonIgnore]
        public long previousID { get; set; }

    }
}