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
using System.Text.RegularExpressions;
using System.Web;
using VisioRestService.Models;

namespace VisioRestService.Manager
{
    public class helper
    {
        List<ExportElement> sorttmp = new List<ExportElement>();
        public void Readfile(string fileName)
        {

            Shape s = new Shape();
            string directoryPath = Path.GetDirectoryName(fileName);
            string docPath = fileName;
            Diagram diagram = new Diagram(docPath);
            int pageCount = 0;
            DataTable dt;
            try
            {
                foreach (Page p in diagram.Pages)
                {

                    ArrayList GridData = new ArrayList();

                    dt = AddColumnsProgrammatically();
                    GetvisioAspose(fileName, pageCount);



                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public DataTable AddColumnsProgrammatically()
        {



            try
            {
                DataTable dt = new DataTable();
                dt.Columns.Add(new System.Data.DataColumn("ShapeID", typeof(Int32)));
                dt.Columns.Add(new System.Data.DataColumn("Text", typeof(string)));
                dt.Columns.Add(new System.Data.DataColumn("elementType", typeof(string)));
                dt.Columns.Add(new System.Data.DataColumn("fromID", typeof(Int32)));
                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }


        public List<ExportElement> GetvisioAspose(string fileName, int pageId)
        {
            ExportElement exportElement;
            List<ExportElement> exportElements = new List<ExportElement>();

            try
            {
                string shapePrefix = "T";
                Shape s = new Shape();
                string directoryPath = Path.GetDirectoryName(fileName);
                string docPath = fileName; //"D:" + @"\visiotoBPMN\BpmnDraw.vsdx";
                Diagram diagram = new Diagram(docPath);



                foreach (Shape shape in diagram.Pages[pageId].Shapes)
                {
                    Shape sh = diagram.Pages[pageId].Shapes.GetShape(shape.ID);
                    exportElement = new ExportElement();

                    if (sh.TwoD && !sh.Name.Contains("CFF Container") && !sh.Name.Contains("Swimlane") && !sh.Name.Contains("Phase List") && !sh.Name.Contains("Separator") && !sh.Name.Contains("On-page reference"))
                    {
                        // Filter shapes by type Foreign
                        if (shape.Type == Aspose.Diagram.TypeValue.Group)
                        {

                            if (shape.NameU != "")
                            {
                                long[] connectedOutShapeIds = shape.ConnectedShapes(ConnectedShapesFlags.ConnectedShapesIncomingNodes, null);
                                if (connectedOutShapeIds.Length == 0)
                                {
                                    exportElement.fromID = shapePrefix + shape.ID;
                                    exportElement.ShapeID = shapePrefix + shape.ID;

                                    if (shape.Text.Value.Count > 1)
                                    {
                                        exportElement.elementName = shape.Text.Value[1].Value.TrimEnd();
                                    }
                                    else if (shape.Text.Value.Count > 0)
                                    {
                                        exportElement.elementName = shape.Text.Value[0].Value.TrimEnd();
                                    }
                                    else { exportElement.elementName = ""; }



                                    if (shape.Acts.Count > 0)
                                    {
                                        if (shape.Acts.Count == 1)
                                        {
                                            if (shape.Acts[0].Name.TrimEnd().Contains("TaskType"))
                                            {
                                                exportElement.elementType = shape.Acts[0].Name.TrimEnd().Substring(8);
                                            }
                                            else if (shape.Acts[0].Name.TrimEnd().Contains("Exclusive"))
                                            {
                                                exportElement.elementType = shape.Acts[0].Name.TrimEnd();
                                            }
                                        }
                                        else
                                        {
                                            if (shape.Acts[1].Name.TrimEnd().Contains("TaskType"))
                                            {
                                                exportElement.elementType = shape.Acts[1].Name.TrimEnd().Substring(8);
                                            }
                                            else
                                            {
                                                exportElement.elementType = shape.Acts[1].Name.TrimEnd();
                                            }
                                        }
                                    }

                                    if (Regex.Replace(shape.Text.Value.Text, "\\<.*?>", "").TrimEnd().Contains("Start"))
                                    {
                                        exportElement.elementType = "start";
                                    }
                                    else if (Regex.Replace(shape.Text.Value.Text, "\\<.*?>", "").TrimEnd().Contains("End"))
                                    {
                                        exportElement.elementType = "end";
                                    }
                                    else
                                    {
                                        exportElement.elementType = shape.Type.ToString();
                                    }


                                    exportElements.Add(exportElement);
                                }
                                else if (connectedOutShapeIds.Length > 1)
                                {
                                    foreach (long l in connectedOutShapeIds)
                                    {
                                        exportElement = new ExportElement();

                                        exportElement.fromID = shapePrefix + l;
                                        exportElement.ShapeID = shapePrefix + shape.ID;
                                        if (shape.Acts.Count > 0)
                                        {
                                            if (shape.Acts.Count == 1)
                                            {
                                                if (shape.Acts[0].Name.TrimEnd().Contains("TaskType"))
                                                {
                                                    exportElement.elementType = shape.Acts[0].Name.TrimEnd().Substring(8);
                                                }
                                                else if (shape.Acts[0].Name.TrimEnd().Contains("Exclusive"))
                                                {
                                                    exportElement.elementType = shape.Acts[0].Name.TrimEnd();
                                                }
                                            }
                                            else
                                            {
                                                if (shape.Acts[1].Name.TrimEnd().Contains("TaskType"))
                                                {
                                                    exportElement.elementType = shape.Acts[1].Name.TrimEnd().Substring(8);
                                                }
                                                else
                                                {
                                                    exportElement.elementType = shape.Acts[1].Name.TrimEnd();
                                                }
                                            }
                                        }

                                        if (shape.Text.Value.Count > 1)
                                        {
                                            exportElement.elementName = shape.Text.Value[1].Value.TrimEnd();
                                        }
                                        else if (shape.Text.Value.Count > 0)
                                        {
                                            exportElement.elementName = shape.Text.Value[0].Value.TrimEnd();
                                        }
                                        else { exportElement.elementName = ""; }
                                        exportElements.Add(exportElement);
                                    }
                                }
                                else
                                {

                                    exportElement.fromID = shapePrefix + connectedOutShapeIds[0];
                                    exportElement.ShapeID = shapePrefix + shape.ID;

                                    if (shape.Text.Value.Count > 1)
                                    {
                                        exportElement.elementName = shape.Text.Value[1].Value.TrimEnd();
                                    }
                                    else if (shape.Text.Value.Count > 0)
                                    {
                                        exportElement.elementName = shape.Text.Value[0].Value.TrimEnd();
                                    }
                                    else { exportElement.elementName = ""; }



                                    if (shape.Acts.Count > 0)
                                    {
                                        if (shape.Acts.Count == 1)
                                        {
                                            if (shape.Acts[0].Name.TrimEnd().Contains("TaskType"))
                                            {
                                                exportElement.elementType = shape.Acts[0].Name.TrimEnd().Substring(8);
                                            }
                                            else if (shape.Acts[0].Name.TrimEnd().Contains("Exclusive"))
                                            {
                                                exportElement.elementType = shape.Acts[0].Name.TrimEnd();
                                            }
                                        }
                                        else
                                        {
                                            if (shape.Acts[1].Name.TrimEnd().Contains("TaskType"))
                                            {
                                                exportElement.elementType = shape.Acts[1].Name.TrimEnd().Substring(8);
                                            }
                                            else
                                            {
                                                exportElement.elementType = shape.Acts[1].Name.TrimEnd();
                                            }
                                        }
                                    }

                                    exportElements.Add(exportElement);
                                }
                            }

                            else
                            {
                                exportElement.elementType = "userTask";
                            }




                        }
                        else if (shape.Type == Aspose.Diagram.TypeValue.Shape)
                        {

                            Regex.Replace(shape.Text.Value.Text, "\\<.*?>", "");
                            //exportElement = new ExportElement();
                            // exportElement.Name = shape.NameU;
                            //exportElement.ShapeID = shape.ID;
                            if (Regex.Replace(shape.Text.Value.Text, "\\<.*?>", "").TrimEnd().Contains("Start"))
                            {
                                exportElement.elementType = "start";
                            }
                            else if (Regex.Replace(shape.Text.Value.Text, "\\<.*?>", "").TrimEnd().Contains("End"))
                            {
                                exportElement.elementType = "end";
                            }
                            else
                            {
                                exportElement.elementType = shape.Type.ToString();
                            }
                            //if (shape.Text.Value.Count == 1)
                            //{
                            //    exportElement.elementName = shape.Text.Value[0].Value.TrimEnd();
                            //}
                            //else if (shape.Text.Value.Count == 2)
                            //{

                            //    exportElement.elementName = shape.Text.Value[1].Value.Trim();
                            //}
                            //else
                            //{
                            //    exportElement.elementName = shape.Master.Name;
                            //}

                            //// exportElement.seqFlows = seqList;

                            //if (exportElement.elementType == "Shape")
                            //    exportElement.elementType = "userTask";
                            //exportElements.Add(exportElement);

                            long[] connectedOutShapeIds = shape.ConnectedShapes(ConnectedShapesFlags.ConnectedShapesIncomingNodes, null);
                            if (connectedOutShapeIds.Length == 0)
                            {
                                exportElement.fromID = shapePrefix + shape.ID;
                                exportElement.ShapeID = shapePrefix + shape.ID;


                                exportElement.elementName = Regex.Replace(shape.Text.Value.Text, "\\<.*?>", "").TrimEnd();
                                //if (shape.Text.Value.Count == 1)
                                //{
                                //    exportElement.elementName = shape.Text.Value[0].Value.TrimEnd();
                                //}
                                //else if (shape.Text.Value.Count == 2)
                                //{

                                //    exportElement.elementName = shape.Text.Value[1].Value.Trim();
                                //}
                                //else
                                //{
                                //    exportElement.elementName = shape.Master.Name;
                                //}

                                // exportElement.seqFlows = seqList;

                                if (exportElement.elementType == "Shape")
                                    exportElement.elementType = "userTask";
                                exportElements.Add(exportElement);


                                // exportElements.Add(exportElement);
                            }
                            else if (connectedOutShapeIds.Length > 1)
                            {
                                foreach (long l in connectedOutShapeIds)
                                {

                                    exportElement = new ExportElement();
                                    if (Regex.Replace(shape.Text.Value.Text, "\\<.*?>", "").TrimEnd().Contains("Start"))
                                    {
                                        exportElement.elementType = "start";
                                    }
                                    else if (Regex.Replace(shape.Text.Value.Text, "\\<.*?>", "").TrimEnd().Contains("End"))
                                    {
                                        exportElement.elementType = "end";
                                    }
                                    else
                                    {
                                        exportElement.elementType = shape.Type.ToString();
                                    }
                                    exportElement.fromID = shapePrefix + l;
                                    exportElement.ShapeID = shapePrefix + shape.ID;
                                    if (exportElement.elementType == "" || exportElement.elementType == null)
                                    {
                                        if (shape.Acts.Count > 0)
                                        {
                                            if (shape.Acts.Count == 1)
                                            {
                                                if (shape.Acts[0].Name.TrimEnd().Contains("TaskType"))
                                                {
                                                    exportElement.elementType = shape.Acts[0].Name.TrimEnd().Substring(8);
                                                }
                                                else if (shape.Acts[0].Name.TrimEnd().Contains("Exclusive"))
                                                {
                                                    exportElement.elementType = shape.Acts[0].Name.TrimEnd();
                                                }
                                            }
                                            else
                                            {
                                                if (shape.Acts[1].Name.TrimEnd().Contains("TaskType"))
                                                {
                                                    exportElement.elementType = shape.Acts[1].Name.TrimEnd().Substring(8);
                                                }
                                                else
                                                {
                                                    exportElement.elementType = shape.Acts[1].Name.TrimEnd();
                                                }
                                            }
                                        }
                                    }

                                    if (shape.Text.Value.Count > 1)
                                    {
                                        exportElement.elementName = shape.Text.Value[1].Value.TrimEnd();
                                    }
                                    else if (shape.Text.Value.Count > 0)
                                    {
                                        exportElement.elementName = shape.Text.Value[0].Value.TrimEnd();
                                    }
                                    else { exportElement.elementName = ""; }
                                    exportElements.Add(exportElement);
                                }
                            }
                            else
                            {

                                exportElement.fromID = shapePrefix + connectedOutShapeIds[0];
                                exportElement.ShapeID = shapePrefix + shape.ID;
                                exportElement.elementName = Regex.Replace(shape.Text.Value.Text, "\\<.*?>", "").TrimEnd();
                                //if (shape.Text.Value.Count > 1)
                                //{
                                //    exportElement.elementName = shape.Text.Value[1].Value.TrimEnd();
                                //}
                                //else if (shape.Text.Value.Count > 0)
                                //{
                                //    exportElement.elementName = shape.Text.Value[0].Value.TrimEnd();
                                //}
                                //else { exportElement.elementName = ""; }

                                if (exportElement.elementType == null)
                                {

                                    if (shape.Acts.Count > 0)
                                    {
                                        if (shape.Acts.Count == 1)
                                        {
                                            if (shape.Acts[0].Name.TrimEnd().Contains("TaskType"))
                                            {
                                                exportElement.elementType = shape.Acts[0].Name.TrimEnd().Substring(8);
                                            }
                                            else if (shape.Acts[0].Name.TrimEnd().Contains("Exclusive"))
                                            {
                                                exportElement.elementType = shape.Acts[0].Name.TrimEnd();
                                            }
                                        }
                                        else
                                        {
                                            if (shape.Acts[1].Name.TrimEnd().Contains("TaskType"))
                                            {
                                                exportElement.elementType = shape.Acts[1].Name.TrimEnd().Substring(8);
                                            }
                                            else
                                            {
                                                exportElement.elementType = shape.Acts[1].Name.TrimEnd();
                                            }
                                        }

                                    }
                                }

                                exportElements.Add(exportElement);
                            }

                        }
                    }
                    if (exportElement.elementType == "Shape" || exportElement.elementType == "User")
                    {
                        exportElement.elementType = "userTask";
                    }
                    if (exportElement.elementType == "Service")
                    {
                        exportElement.elementType = "serviceTask";
                    }
                    if (exportElement.elementType == null)
                    {
                        exportElement.elementType = "userTask";
                    }
                    if (exportElement.elementType.Contains("Exclusive") || exportElement.elementType.Contains("Gateway"))
                    { exportElement.elementType = "exclusiveGateway"; }
                }


                return exportElements;

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        public List<ExportElement> GetvisioAsposetest(string fileName, int pageId)
        {
            ExportElement exportElement;
            List<ExportElement> exportElements = new List<ExportElement>();

            try
            {
                string shapePrefix = "T";
                Shape s = new Shape();
                string directoryPath = Path.GetDirectoryName(fileName);
                string docPath = fileName; //"D:" + @"\visiotoBPMN\BpmnDraw.vsdx";
                Diagram diagram = new Diagram(docPath);



                foreach (Shape shape in diagram.Pages[pageId].Shapes)
                {
                    Shape sh = diagram.Pages[pageId].Shapes.GetShape(shape.ID);
                    exportElement = new ExportElement();

                    #region Main shape
                    if (sh.TwoD && !sh.Name.Contains("CFF Container") && !sh.Name.Contains("Swimlane") && !sh.Name.Contains("Phase List") && !sh.Name.Contains("Separator") && !sh.Name.Contains("Manual") && !sh.Name.ToLower().Contains("on-page reference"))
                    {
                        if (sh.Text.Value.Text != string.Empty)
                        {
                            if (sh.Name.ToLower().Contains("annotation"))
                            {
                                //exportElement.elementName = Regex.Replace(sh.Text.Value.Text, "\\<.*?>", "").TrimEnd();
                                //exportElement.elementType = "textAnnotation";
                                //exportElement.ShapeID = shapePrefix + sh.ID;
                                //exportElement.fromID = sh.GluedShapes(GluedShapesFlags.GluedShapesOutgoing1D, null, null).ToString();

                                //exportElement.elementName = Regex.Replace(sh.Text.Value.Text, "\\<.*?>", "").TrimEnd();
                                //exportElement.elementType = "boundaryEvent";
                                //exportElement.ShapeID = sh.ID;
                                //if (sh.Geoms.Count > 0)
                                //{
                                //    exportElement.fromID = sh.Geoms[0].NextCoordinateIX;
                                //}
                            }
                            else
                            {
                                exportElement.elementName = Regex.Replace(sh.Text.Value.Text, "\\<.*?>", "").TrimEnd();
                                exportElement.elementType = (sh.Type.ToString() == "Shape") ? "userTask" : sh.Type.ToString();

                                if (sh.Name.ToLower().Contains("decision") || sh.Name.ToLower().Contains("gateway"))
                                {
                                    exportElement.elementType = "exclusiveGateway";
                                }

                                long[] connectedOutShapeIds = sh.ConnectedShapes(ConnectedShapesFlags.ConnectedShapesIncomingNodes, null);
                                if (connectedOutShapeIds.Length == 0)
                                {
                                    exportElement.ShapeID = shapePrefix + sh.ID;
                                    exportElement.fromID = shapePrefix + shape.ID;
                                    exportElement.SeqID = shape.ID;
                                    exportElements.Add(exportElement);
                                }
                                else if (connectedOutShapeIds.Length > 1)
                                {
                                    foreach (long l in connectedOutShapeIds)
                                    {
                                        exportElement = new ExportElement();
                                        exportElement.elementName = Regex.Replace(sh.Text.Value.Text, "\\<.*?>", "").TrimEnd();
                                        exportElement.elementType = (sh.Type.ToString() == "Shape") ? "userTask" : sh.Type.ToString();
                                        exportElement.fromID = shapePrefix + l;
                                        exportElement.ShapeID = shapePrefix + shape.ID;
                                        exportElement.SeqID = shape.ID;
                                        if (exportElement.elementName != null)
                                        {
                                            if (exportElement.elementName.ToUpper().Contains("START"))
                                            {
                                                exportElement.elementType = "start";
                                            }
                                            if (exportElement.elementName.ToUpper().Contains("END"))
                                            {
                                                exportElement.elementType = "end";
                                            }

                                        }
                                        if (sh.Name.ToLower().Contains("decision") || sh.Name.ToLower().Contains("gateway"))
                                        {
                                            exportElement.elementType = "exclusiveGateway";
                                        }
                                        exportElements.Add(exportElement);
                                    }
                                }
                                else
                                {
                                    exportElement.fromID = shapePrefix + connectedOutShapeIds[0];
                                    exportElement.ShapeID = shapePrefix + shape.ID;
                                    exportElement.SeqID = shape.ID;
                                    exportElements.Add(exportElement);
                                }

                            }

                        }
                        #region Shapes inside shape
                        else if (sh.Shapes.Count > 0)
                        {
                            foreach (Shape sh1 in sh.Shapes)
                            {
                                if (sh1.Text.Value.Text == string.Empty)
                                {
                                    if (sh1.Shapes.Count > 0)
                                    {
                                        foreach (Shape sh2 in sh1.Shapes)
                                        {

                                            if (sh2.Shapes.Count > 0)
                                            {
                                                foreach (Shape sh3 in sh2.Shapes)
                                                {
                                                    exportElement = new ExportElement();
                                                    exportElement.elementName = Regex.Replace(sh3.Text.Value.Text, "\\<.*?>", "").TrimEnd();
                                                    exportElement.elementType = (sh3.Type.ToString() == "Shape") ? "userTask" : sh3.Type.ToString();


                                                    long[] connectedOutShapeIds = sh3.ConnectedShapes(ConnectedShapesFlags.ConnectedShapesIncomingNodes, null);
                                                    if (connectedOutShapeIds.Length == 0)
                                                    {
                                                        exportElement.ShapeID = shapePrefix + sh3.ID;
                                                        exportElement.fromID = shapePrefix + sh3.ID;
                                                        exportElement.SeqID = sh3.ID;
                                                    }
                                                    else if (connectedOutShapeIds.Length > 1)
                                                    {
                                                        foreach (long l in connectedOutShapeIds)
                                                        {
                                                            exportElement = new ExportElement();
                                                            exportElement.elementName = Regex.Replace(sh3.Text.Value.Text, "\\<.*?>", "").TrimEnd();
                                                            exportElement.elementType = (sh3.Type.ToString() == "Shape") ? "userTask" : sh3.Type.ToString();
                                                            exportElement.fromID = shapePrefix + l;
                                                            exportElement.ShapeID = shapePrefix + sh3.ID;
                                                            exportElement.SeqID = sh3.ID;
                                                            exportElements.Add(exportElement);
                                                        }
                                                    }
                                                    else
                                                    {
                                                        exportElement.fromID = shapePrefix + connectedOutShapeIds[0];
                                                        exportElement.ShapeID = shapePrefix + sh3.ID;
                                                        exportElement.SeqID = sh3.ID;
                                                    }
                                                    exportElements.Add(exportElement);
                                                }
                                            }
                                            else
                                            {
                                                exportElement = new ExportElement();
                                                exportElement.elementName = Regex.Replace(sh2.Text.Value.Text, "\\<.*?>", "").TrimEnd();
                                                exportElement.elementType = (sh2.Type.ToString() == "Shape") ? "userTask" : sh2.Type.ToString();


                                                long[] connectedOutShapeIds = sh2.ConnectedShapes(ConnectedShapesFlags.ConnectedShapesIncomingNodes, null);
                                                if (connectedOutShapeIds.Length == 0)
                                                {
                                                    exportElement.ShapeID = shapePrefix + sh2.ID;
                                                    exportElement.fromID = shapePrefix + sh2.ID;
                                                    exportElement.SeqID = sh2.ID;
                                                }
                                                else if (connectedOutShapeIds.Length > 1)
                                                {
                                                    foreach (long l in connectedOutShapeIds)
                                                    {
                                                        exportElement = new ExportElement();
                                                        exportElement.elementName = Regex.Replace(sh2.Text.Value.Text, "\\<.*?>", "").TrimEnd();
                                                        exportElement.elementType = (sh2.Type.ToString() == "Shape") ? "userTask" : sh2.Type.ToString();
                                                        exportElement.fromID = shapePrefix + l;
                                                        exportElement.ShapeID = shapePrefix + sh2.ID;
                                                        exportElement.SeqID = sh2.ID;
                                                        exportElements.Add(exportElement);
                                                    }
                                                }
                                                else
                                                {
                                                    exportElement.fromID = shapePrefix + connectedOutShapeIds[0];
                                                    exportElement.ShapeID = shapePrefix + sh2.ID;
                                                    exportElement.SeqID = sh2.ID;
                                                }
                                                exportElements.Add(exportElement);
                                            }
                                        }

                                    }


                                }
                                else
                                {
                                    exportElement = new ExportElement();
                                    exportElement.elementName = Regex.Replace(sh1.Text.Value.Text, "\\<.*?>", "").TrimEnd();
                                    exportElement.elementType = (sh1.Type.ToString() == "Shape") ? "userTask" : sh1.Type.ToString();

                                    long[] connectedOutShapeIds = sh1.ConnectedShapes(ConnectedShapesFlags.ConnectedShapesIncomingNodes, null);
                                    if (connectedOutShapeIds.Length == 0)
                                    {
                                        exportElement.ShapeID = shapePrefix + sh1.ID;
                                        exportElement.fromID = shapePrefix + sh1.ID;
                                        exportElement.SeqID = sh1.ID;
                                        exportElements.Add(exportElement);
                                    }
                                    else if (connectedOutShapeIds.Length > 1)
                                    {
                                        foreach (long l in connectedOutShapeIds)
                                        {
                                            exportElement = new ExportElement();
                                            exportElement.elementName = Regex.Replace(sh1.Text.Value.Text, "\\<.*?>", "").TrimEnd();
                                            exportElement.elementType = (sh1.Type.ToString() == "Shape") ? "userTask" : sh1.Type.ToString();
                                            exportElement.fromID = shapePrefix + l;
                                            exportElement.ShapeID = shapePrefix + sh1.ID;
                                            exportElement.SeqID = sh1.ID;

                                            exportElements.Add(exportElement);
                                        }
                                    }
                                    else
                                    {
                                        exportElement.fromID = shapePrefix + connectedOutShapeIds[0];
                                        exportElement.ShapeID = shapePrefix + sh1.ID;
                                        exportElement.SeqID = sh1.ID;
                                        exportElements.Add(exportElement);
                                    }

                                }
                            }
                        }
                        #endregion
                        else
                        {
                            if (!string.IsNullOrEmpty(exportElement.elementName))
                            {
                                exportElement.elementName = Regex.Replace(sh.Text.Value.Text, "\\<.*?>", "").TrimEnd();
                                exportElement.elementType = (sh.Type.ToString() == "Shape") ? "userTask" : sh.Type.ToString();

                                long[] connectedOutShapeIds = shape.ConnectedShapes(ConnectedShapesFlags.ConnectedShapesIncomingNodes, null);
                                if (connectedOutShapeIds.Length == 0)
                                {
                                    exportElement.ShapeID = shapePrefix + sh.ID;
                                    exportElement.fromID = shapePrefix + shape.ID;
                                    exportElement.SeqID = shape.ID;
                                    exportElements.Add(exportElement);
                                }
                                else if (connectedOutShapeIds.Length > 1)
                                {
                                    foreach (long l in connectedOutShapeIds)
                                    {
                                        exportElement = new ExportElement();
                                        exportElement.elementName = Regex.Replace(sh.Text.Value.Text, "\\<.*?>", "").TrimEnd();
                                        exportElement.elementType = sh.Type.ToString();
                                        exportElement.fromID = shapePrefix + l;
                                        exportElement.ShapeID = shapePrefix + shape.ID;
                                        exportElement.SeqID = shape.ID;
                                        exportElements.Add(exportElement);
                                    }
                                }
                                else
                                {
                                    exportElement.fromID = shapePrefix + connectedOutShapeIds[0];
                                    exportElement.ShapeID = shapePrefix + shape.ID;
                                    exportElement.SeqID = shape.ID;
                                    exportElements.Add(exportElement);
                                }
                            }
                            else
                            {
                                exportElement.elementName = Regex.Replace(sh.Name, "\\<.*?>", "").TrimEnd();
                                exportElement.elementType = (sh.Type.ToString() == "Shape") ? "userTask" : sh.Type.ToString();
                                exportElement.fromID = shapePrefix + shape.ID;
                                exportElement.ShapeID = shapePrefix + shape.ID;
                                exportElement.SeqID = shape.ID;
                                exportElements.Add(exportElement);
                            }
                        }
                        // exportElements.Add(exportElement);
                    }
                    #endregion


                    if (exportElement.elementName != null)
                    {
                        if (exportElement.elementName.ToUpper().Contains("START"))
                        {
                            exportElement.elementType = "start";
                        }
                        if (exportElement.elementName.ToUpper().Contains("END"))
                        {
                            exportElement.elementType = "end";
                        }

                    }
                    if (exportElement.elementType == "Shape" || exportElement.elementType == "User")
                    {
                        exportElement.elementType = "userTask";
                    }
                    if (exportElement.elementType == "Service")
                    {
                        exportElement.elementType = "serviceTask";
                    }
                    if (exportElement.elementType == null)
                    {
                        exportElement.elementType = "userTask";
                    }
                    if (exportElement.elementType == "Group")
                    {
                        exportElement.elementType = "userTask";
                    }
                    if (exportElement.elementType.Contains("Exclusive") || exportElement.elementType.Contains("Gateway"))
                    { exportElement.elementType = "exclusiveGateway"; }

                }



                return exportElements.OrderBy(x => x.SeqID).ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private List<ExportElement> Rearrrange(List<ExportElement> exp)
        {
            List<ExportElement> tmp = new List<ExportElement>();

            //List<ExportElement> sorttmp = new List<ExportElement>();



            if (exp.Count > 0)
            {

                tmp = Reloop(exp);
                while (tmp.Count != 0)
                {
                    tmp = Reloop(tmp);
                }
            }


            //for (int i = 0; i <= exp.Count - 1; i++)
            //{


            //    if (sorttmp.Count == 0)
            //    {
            //        sorttmp.Add(exp[i]);
            //        exp.Remove(exp[i]);
            //    }
            //    else
            //    {
            //        bool target = sorttmp.Exists(item => item.SeqID.ToString().Contains(exp[i].previousID.ToString()));
            //        if (target == true)
            //        {
            //            sorttmp.Add(exp[i]);
            //            exp.Remove(exp[i]);
            //        }
            //        //else
            //        //{
            //        //    tmp.Add(exp[i]);
            //        //}
            //    }


            //}



            return sorttmp;

        }
        private List<ExportElement> Reloop(List<ExportElement> exp)
        {
            List<ExportElement> tmp = new List<ExportElement>();

            try
            {

                if (exp.Count > 0)
                {
                    for (int i = 0; i <= exp.Count - 1; i++)
                    {


                        if (sorttmp.Count == 0)
                        {
                            sorttmp.Add(exp[i]);

                        }
                        else
                        {
                            if (exp[i].SeqID.ToString() == exp[i].previousID.ToString() && !sorttmp.Exists(item => item.SeqID.ToString().Contains(exp[i].SeqID.ToString())))
                            {
                                sorttmp.Add(exp[i]);
                            }
                            else
                            {
                                bool target = sorttmp.Exists(item => item.SeqID.ToString().Contains(exp[i].previousID.ToString()));
                                if (target == true)
                                {
                                    sorttmp.Add(exp[i]);

                                }
                                else
                                {
                                    tmp.Add(exp[i]);
                                }
                            }
                        }


                    }

                }
            }
            catch (Exception ex)
            {

            }
            return tmp;
        }
        public void getAllshapes(string fileName, int pageId)
        {
            List<ExportElement> exportElements = new List<ExportElement>();
            string shapePrefix = "T";
            Shape s = new Shape();
            string directoryPath = Path.GetDirectoryName(fileName);
            string docPath = fileName; //"D:" + @"\visiotoBPMN\BpmnDraw.vsdx";
            Diagram diagram = new Diagram(docPath);

            foreach (Shape shape in diagram.Pages[pageId].Shapes)
            {
                ExportElement exportElement = new ExportElement();
                
                exportElement.ShapeID = shapePrefix + shape.ID;
                exportElement.fromID = shapePrefix + shape.ID;
                exportElement.SeqID = shape.ID;
                exportElement.previousID = shape.ID;
                exportElement.elementName = shape.Name;
                exportElements.Add(exportElement);
                exportElements.OrderBy(x => x.SeqID).ToList();
            }
        }
        public List<ExportElement> GetvisioAsposeSingleshape(string fileName, int pageId)
        {
            ExportElement exportElement;
            List<ExportElement> exportElements = new List<ExportElement>();

            try
            {
                string shapePrefix = "T";
                Shape s = new Shape();
                string directoryPath = Path.GetDirectoryName(fileName);
                string docPath = fileName; //"D:" + @"\visiotoBPMN\BpmnDraw.vsdx";
                Diagram diagram = new Diagram(docPath);



                foreach (Shape shape in diagram.Pages[pageId].Shapes)
                {
                    Shape sh = diagram.Pages[pageId].Shapes.GetShape(shape.ID);
                    exportElement = new ExportElement();

                    //&& !sh.Name.ToLower().Contains("on-page reference")
                    #region Main shape 
                    if (sh.TwoD && !sh.Name.Contains("CFF Container") && !sh.Name.Contains("Swimlane") && !sh.Name.Contains("Phase List") && !sh.Name.Contains("Separator") && !sh.Name.Contains("Manual") && !sh.Name.ToLower().Contains("on-page reference"))
                    {
                        if (sh.Text.Value.Text != string.Empty)
                        {
                            if (sh.Name.ToLower().Contains("annotation"))
                            {
                                //exportElement.elementName = Regex.Replace(sh.Text.Value.Text, "\\<.*?>", "").TrimEnd();
                                //exportElement.elementType = "textAnnotation";
                                //exportElement.ShapeID = shapePrefix + sh.ID;
                                //exportElement.fromID = sh.GluedShapes(GluedShapesFlags.GluedShapesOutgoing1D, null, null).ToString();

                                //exportElement.elementName = Regex.Replace(sh.Text.Value.Text, "\\<.*?>", "").TrimEnd();
                                //exportElement.elementType = "boundaryEvent";
                                //exportElement.ShapeID = sh.ID;
                                //if (sh.Geoms.Count > 0)
                                //{
                                //    exportElement.fromID = sh.Geoms[0].NextCoordinateIX;
                                //}
                            }
                            else
                            {
                                exportElement.elementName = Regex.Replace(sh.Text.Value.Text, "\\<.*?>", "").TrimEnd();
                                exportElement.elementType = (sh.Type.ToString() == "Shape") ? "userTask" : sh.Type.ToString();

                                if (sh.Name.ToLower().Contains("decision") || sh.Name.ToLower().Contains("gateway"))
                                {
                                    exportElement.elementType = "exclusiveGateway";
                                }

                                long[] connectedOutShapeIds = sh.ConnectedShapes(ConnectedShapesFlags.ConnectedShapesIncomingNodes, null);
                                if (connectedOutShapeIds.Length == 0)
                                {
                                    exportElement.ShapeID = shapePrefix + sh.ID;
                                    exportElement.fromID = shapePrefix + shape.ID;
                                    exportElement.SeqID = shape.ID;
                                    exportElement.previousID = shape.ID;
                                    exportElements.Add(exportElement);
                                }
                                else if (connectedOutShapeIds.Length > 1)
                                {
                                    foreach (long l in connectedOutShapeIds)
                                    {
                                        exportElement = new ExportElement();
                                        exportElement.elementName = Regex.Replace(sh.Text.Value.Text, "\\<.*?>", "").TrimEnd();
                                        exportElement.elementType = (sh.Type.ToString() == "Shape") ? "userTask" : sh.Type.ToString();
                                        exportElement.fromID = shapePrefix + l;
                                        exportElement.ShapeID = shapePrefix + shape.ID;
                                        exportElement.SeqID = shape.ID;
                                        exportElement.previousID = l;
                                        if (exportElement.elementName != null)
                                        {
                                            if (exportElement.elementName.ToUpper().Contains("START"))
                                            {
                                                exportElement.elementType = "start";
                                            }
                                            if (exportElement.elementName.ToUpper().Contains("END"))
                                            {
                                                exportElement.elementType = "end";
                                            }

                                        }
                                        if (sh.Name.ToLower().Contains("decision") || sh.Name.ToLower().Contains("gateway"))
                                        {
                                            exportElement.elementType = "exclusiveGateway";
                                        }
                                        exportElements.Add(exportElement);
                                    }
                                }
                                else
                                {
                                    exportElement.fromID = shapePrefix + connectedOutShapeIds[0];
                                    exportElement.ShapeID = shapePrefix + shape.ID;
                                    exportElement.SeqID = shape.ID;
                                    exportElement.previousID = connectedOutShapeIds[0];
                                    exportElements.Add(exportElement);
                                }

                            }

                        }
                        #region Shapes inside shape
                        else if (sh.Shapes.Count > 0)
                        {
                            foreach (Shape sh1 in sh.Shapes)
                            {
                                if (sh1.Text.Value.Text == string.Empty)
                                {
                                    if (sh1.Shapes.Count > 0)
                                    {
                                        if (!sh1.Name.ToLower().Contains("On-page reference"))
                                        {
                                            foreach (Shape sh2 in sh1.Shapes)
                                            {


                                                if (sh2.Shapes.Count > 0)
                                                {
                                                    for (int i = 1; i <= sh2.Shapes.Count - 1; i++)
                                                    {

                                                        Shape sh3 = sh2.Shapes[i];
                                                        exportElement = new ExportElement();
                                                        exportElement.elementName = Regex.Replace(sh3.Text.Value.Text, "\\<.*?>", "").TrimEnd();
                                                        exportElement.elementType = (sh3.Type.ToString() == "Shape") ? "userTask" : sh3.Type.ToString();

                                                        long[] connectedOutShapeIds = sh3.ConnectedShapes(ConnectedShapesFlags.ConnectedShapesIncomingNodes, null);
                                                        if (connectedOutShapeIds.Length == 0)
                                                        {
                                                            exportElement.ShapeID = shapePrefix + sh3.ID;
                                                            exportElement.fromID = shapePrefix + sh3.ID;
                                                            exportElement.SeqID = sh3.ID;
                                                            exportElement.previousID = sh3.ID;
                                                        }
                                                        else if (connectedOutShapeIds.Length > 1)
                                                        {
                                                            foreach (long l in connectedOutShapeIds)
                                                            {
                                                                exportElement = new ExportElement();
                                                                exportElement.elementName = Regex.Replace(sh3.Text.Value.Text, "\\<.*?>", "").TrimEnd();
                                                                exportElement.elementType = (sh3.Type.ToString() == "Shape") ? "userTask" : sh3.Type.ToString();
                                                                exportElement.fromID = shapePrefix + l;
                                                                exportElement.ShapeID = shapePrefix + sh3.ID;
                                                                exportElement.SeqID = sh3.ID;
                                                                exportElement.previousID = l;
                                                                exportElements.Add(exportElement);
                                                            }
                                                        }
                                                        else
                                                        {
                                                            exportElement.fromID = shapePrefix + connectedOutShapeIds[0];
                                                            exportElement.ShapeID = shapePrefix + sh3.ID;
                                                            exportElement.SeqID = sh3.ID;
                                                            exportElement.previousID = connectedOutShapeIds[0];
                                                        }
                                                        exportElements.Add(exportElement);
                                                    }
                                                }
                                                else
                                                {
                                                    if (sh2.Fill.FillForegnd.Value != "#ff9900") 
                                                    {

                                                        exportElement = new ExportElement();
                                                        exportElement.elementName = Regex.Replace(sh2.Text.Value.Text, "\\<.*?>", "").TrimEnd();
                                                        exportElement.elementType = (sh2.Type.ToString() == "Shape") ? "userTask" : sh2.Type.ToString();

                                                        long[] connectedOutShapeIds = sh2.ConnectedShapes(ConnectedShapesFlags.ConnectedShapesIncomingNodes, null);
                                                        if (connectedOutShapeIds.Length == 0)
                                                        {
                                                            exportElement.ShapeID = shapePrefix + sh2.ID;
                                                            exportElement.fromID = shapePrefix + sh2.ID;
                                                            exportElement.SeqID = sh2.ID;
                                                            exportElement.previousID = sh2.ID;
                                                            exportElements.Add(exportElement);
                                                        }
                                                        else if (connectedOutShapeIds.Length > 1)
                                                        {
                                                            foreach (long l in connectedOutShapeIds)
                                                            {
                                                                exportElement = new ExportElement();
                                                                exportElement.elementName = Regex.Replace(sh2.Text.Value.Text, "\\<.*?>", "").TrimEnd();
                                                                exportElement.elementType = (sh2.Type.ToString() == "Shape") ? "userTask" : sh2.Type.ToString();
                                                                exportElement.fromID = shapePrefix + l;
                                                                exportElement.ShapeID = shapePrefix + sh2.ID;
                                                                exportElement.SeqID = sh2.ID;
                                                                exportElement.previousID = l;
                                                                exportElements.Add(exportElement);
                                                            }
                                                        }
                                                        else
                                                        {
                                                            exportElement.fromID = shapePrefix + connectedOutShapeIds[0];
                                                            exportElement.ShapeID = shapePrefix + sh2.ID;
                                                            exportElement.SeqID = sh2.ID;
                                                            exportElement.previousID = connectedOutShapeIds[0];
                                                            exportElements.Add(exportElement);
                                                        }

                                                    }
                                                }
                                            }
                                        }

                                    }


                                }
                                else
                                {
                                    if (!sh1.Name.ToLower().Contains("on-page reference"))
                                    {
                                        if (sh1.Fill.FillForegnd.Value != "#ff9900")
                                        {
                                            exportElement = new ExportElement();
                                            exportElement.elementName = Regex.Replace(sh1.Text.Value.Text, "\\<.*?>", "").TrimEnd();
                                            exportElement.elementType = (sh1.Type.ToString() == "Shape") ? "userTask" : sh1.Type.ToString();

                                            long[] connectedOutShapeIds = sh1.ConnectedShapes(ConnectedShapesFlags.ConnectedShapesIncomingNodes, null);
                                            if (connectedOutShapeIds.Length == 0)
                                            {
                                                exportElement.ShapeID = shapePrefix + sh1.ID;
                                                exportElement.fromID = shapePrefix + sh1.ID;
                                                exportElement.SeqID = sh1.ID;
                                                exportElement.previousID = sh1.ID;
                                                exportElements.Add(exportElement);
                                            }
                                            else if (connectedOutShapeIds.Length > 1)
                                            {
                                                foreach (long l in connectedOutShapeIds)
                                                {
                                                    exportElement = new ExportElement();
                                                    exportElement.elementName = Regex.Replace(sh1.Text.Value.Text, "\\<.*?>", "").TrimEnd();
                                                    exportElement.elementType = (sh1.Type.ToString() == "Shape") ? "userTask" : sh1.Type.ToString();
                                                    exportElement.fromID = shapePrefix + l;
                                                    exportElement.ShapeID = shapePrefix + sh1.ID;
                                                    exportElement.SeqID = sh1.ID;
                                                    exportElement.previousID = l;
                                                    exportElements.Add(exportElement);
                                                }
                                            }
                                            else
                                            {
                                                exportElement.fromID = shapePrefix + connectedOutShapeIds[0];
                                                exportElement.ShapeID = shapePrefix + sh1.ID;
                                                exportElement.SeqID = sh1.ID;
                                                exportElement.previousID = connectedOutShapeIds[0];
                                                exportElements.Add(exportElement);
                                            }
                                        }
                                    }

                                }
                            }
                        }
                        #endregion
                        else
                        {
                            if (!string.IsNullOrEmpty(exportElement.elementName))
                            {
                                exportElement.elementName = Regex.Replace(sh.Text.Value.Text, "\\<.*?>", "").TrimEnd();
                                exportElement.elementType = (sh.Type.ToString() == "Shape") ? "userTask" : sh.Type.ToString();

                                long[] connectedOutShapeIds = shape.ConnectedShapes(ConnectedShapesFlags.ConnectedShapesIncomingNodes, null);
                                if (connectedOutShapeIds.Length == 0)
                                {
                                    exportElement.ShapeID = shapePrefix + sh.ID;
                                    exportElement.fromID = shapePrefix + shape.ID;
                                    exportElement.SeqID = shape.ID;
                                    exportElement.previousID = shape.ID;
                                    exportElements.Add(exportElement);
                                }
                                else if (connectedOutShapeIds.Length > 1)
                                {
                                    foreach (long l in connectedOutShapeIds)
                                    {
                                        exportElement = new ExportElement();
                                        exportElement.elementName = Regex.Replace(sh.Text.Value.Text, "\\<.*?>", "").TrimEnd();
                                        exportElement.elementType = sh.Type.ToString();
                                        exportElement.fromID = shapePrefix + l;
                                        exportElement.ShapeID = shapePrefix + shape.ID;
                                        exportElement.SeqID = shape.ID;
                                        exportElement.previousID = l;
                                        exportElements.Add(exportElement);
                                    }
                                }
                                else
                                {
                                    exportElement.fromID = shapePrefix + connectedOutShapeIds[0];
                                    exportElement.ShapeID = shapePrefix + shape.ID;
                                    exportElement.SeqID = shape.ID;
                                    exportElement.previousID = connectedOutShapeIds[0];
                                    exportElements.Add(exportElement);
                                }
                            }
                            else
                            {
                                if (sh.Text.Value == null || sh.Text.Value.ToString() == string.Empty)
                                {
                                    exportElement.elementName = Regex.Replace(sh.Name, "\\<.*?>", "").TrimEnd();
                                    exportElement.elementType = (sh.Type.ToString() == "Shape") ? "userTask" : sh.Type.ToString();
                                    exportElement.fromID = shapePrefix + shape.ID;
                                    exportElement.ShapeID = shapePrefix + shape.ID;
                                    exportElement.SeqID = shape.ID;
                                    exportElement.previousID = shape.ID;
                                    exportElements.Add(exportElement);
                                }
                            }
                        }
                        // exportElements.Add(exportElement);
                    }
                    #endregion


                    if (exportElement.elementName != null)
                    {
                        if (exportElement.elementName.ToUpper().Contains("START"))
                        {
                            exportElement.elementType = "start";
                        }
                        if (exportElement.elementName.ToUpper().Contains("END"))
                        {
                            exportElement.elementType = "end";
                        }

                    }
                    if (exportElement.elementType == "Shape" || exportElement.elementType == "User")
                    {
                        exportElement.elementType = "userTask";
                    }
                    if (exportElement.elementType == "Service")
                    {
                        exportElement.elementType = "serviceTask";
                    }
                    if (exportElement.elementType == null)
                    {
                        exportElement.elementType = "userTask";
                    }
                    if (exportElement.elementType == "Group")
                    {
                        exportElement.elementType = "userTask";
                    }
                    if (exportElement.elementType.Contains("Exclusive") || exportElement.elementType.Contains("Gateway"))
                    { exportElement.elementType = "exclusiveGateway"; }

                }

                //return Rearrrange(exportElements.OrderBy(x => x.SeqID).ToList());
                return exportElements.OrderBy(x => x.SeqID).ToList();

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public string GetBPMNXML(string ProcessName, List<ExportElement> elemtDet)
        {
            var client = new RestClient("http://10.10.233.48:9009/mud/postprocess");

            //client.Proxy = new WebProxy("10.254.3.121");
            IWebProxy proxy = WebRequest.GetSystemWebProxy();
            //NetworkCredential nc = new NetworkCredential();
            //CredentialCache cc = new CredentialCache();
            //nc.UserName = "gs00334809";
            //nc.Password = "apr@2019";
            //nc.Domain = "Techmahindra";
            //cc.Add("http://10.254.3.121", 8080, "Basic", nc);
            //proxy.Credentials = cc;

            //client.Proxy.Credentials = proxy.Credentials;

            var request = new RestRequest(Method.POST);


            request.AddHeader("X-Token-Key", "dsds-sdsdsds-swrwerfd-dfdfd");
            request.AddHeader("Content-Type", "application/json");
            var p = new PostProcess { processName = ProcessName, elementDetails = elemtDet };
            // Json to post.
            string jsonToSend = JsonConvert.SerializeObject(p);


            request.RequestFormat = DataFormat.Xml;
            request.AddJsonBody(jsonToSend);

            try
            {
                IRestResponse response = client.Execute(request);
                var content = response.Content; // raw content as string
                return content;
            }
            catch (Exception error)
            {
                return "";
                // Log
            }
        }
    }
}