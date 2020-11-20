using Aspose.Diagram;
using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Web.Http;
using VisioRestService.Models;
using VisioRestService.Manager;
using System.Configuration;

namespace VisioRestService.Controllers
{
    [RoutePrefix("api/Visio")]
    public class VisioController : ApiController
    {
        //[Route("getfile")]
        //public string getfile()
        //{
        //    return "Welcome";
        //}

        [Route("getBpmnfile")]
        [HttpGet]
        public string getBpmnfile(string filename)
        {
            helper obj = new helper();
            if (filename != "")
            {
                string localFilePath = ConfigurationManager.AppSettings["VisioFilePath"];

                string visoFile = "D:\\visiotoBPMN\\"+ filename;
                visoFile = localFilePath + filename;
                var eldata = new List<ExportElement>();
                string strReturn = "";
                if (!File.Exists(visoFile))
                {
                    return "The file not exists.";
                }
                else
                {
                   // obj.getAllshapes(visoFile, 0);
                    eldata = obj.GetvisioAsposeSingleshape(visoFile, 0);

                    strReturn = obj.GetBPMNXML("test", eldata);
                    return strReturn;
                }
            }
            else {
                return "Please Provide File.";
            }
            
        }

       
    }
}
