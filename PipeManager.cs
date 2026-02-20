using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace LogiK3D.Piping
{
    public static class PipeManager
    {
        private static HashSet<ObjectId> _dirtyPolylines = new HashSet<ObjectId>();
        private static bool _isUpdating = false;
        
        public const string LineDataAppName = "LogiK_LineData";
        public const string SolidDataAppName = "LogiK_Data";
        public const string DictName = "LogiK_PipingDict";

        public static void StartReactor()
        {
            Application.DocumentManager.DocumentCreated += (s, e) => {
                e.Document.Database.ObjectModified += Database_ObjectModified;
            };
            
            foreach (Document doc in Application.DocumentManager)
            {
                doc.Database.ObjectModified += Database_ObjectModified;
            }
        }

        private static void Database_ObjectModified(object sender, ObjectEventArgs e)
        {
            if (_isUpdating) return;
            
            if (e.DBObject is Polyline || e.DBObject is Polyline3d)
            {
                using (ResultBuffer rb = e.DBObject.GetXDataForApplication(LineDataAppName))
                {
                    if (rb != null)
                    {
                        _dirtyPolylines.Add(e.DBObject.ObjectId);
                        Application.Idle -= Application_Idle;
                        Application.Idle += Application_Idle;
                    }
                }
            }
        }

        private static void Application_Idle(object sender, EventArgs e)
        {
            Application.Idle -= Application_Idle;
            if (_dirtyPolylines.Count == 0) return;

            _isUpdating = true;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc != null)
            {
                using (DocumentLock docLock = doc.LockDocument())
                {
                    using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                    {
                        foreach (ObjectId polyId in _dirtyPolylines)
                        {
                            if (polyId.IsErased || !polyId.IsValid) continue;
                            
                            Entity poly = tr.GetObject(polyId, OpenMode.ForWrite) as Entity;
                            if (poly != null)
                            {
                                string lineNumber = "";
                                string dn = "";
                                double od = 0;
                                double thickness = 3.2; // Default
                                
                                ResultBuffer rb = poly.GetXDataForApplication(LineDataAppName);
                                if (rb != null)
                                {
                                    TypedValue[] values = rb.AsArray();
                                    if (values.Length >= 4)
                                    {
                                        lineNumber = values[1].Value.ToString();
                                        dn = values[2].Value.ToString();
                                        od = (double)values[3].Value;
                                    }
                                    if (values.Length >= 5)
                                    {
                                        thickness = (double)values[4].Value;
                                    }
                                }
                                
                                if (!string.IsNullOrEmpty(lineNumber) && od > 0)
                                {
                                    GeneratePiping(poly, lineNumber, dn, od, thickness, tr);
                                }
                            }
                        }
                        tr.Commit();
                    }
                }
            }
            _dirtyPolylines.Clear();
            _isUpdating = false;
        }

        private class InlineComponent
        {
            public double Param1 { get; set; }
            public double Param2 { get; set; }
            public Point3d Pt1 { get; set; }
            public Point3d Pt2 { get; set; }
            public string CompType { get; set; }
            public string TargetDN { get; set; }
            public double TargetOD { get; set; }
        }

        public static void GeneratePiping(Entity poly, string lineNumber, string dn, double od, double thickness, Transaction tr)
        {
            Database db = poly.Database;
            
            // 1. Supprimer les anciens solides liés à cette polyligne
            DeleteLinkedSolids(poly, tr);

            // 2. Générer les nouveaux solides
            List<Point3d> vertices = GetPolylineVertices(poly, tr);
            if (vertices.Count < 2) return;

            double pipeRadius = od / 2.0;
            string dnString = dn.Replace("DN", "");
            double dnValue = 100.0;
            double.TryParse(dnString, out dnValue);
            double elbowRadiusLR = 1.5 * dnValue;

            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
            List<Handle> newSolidHandles = new List<Handle>();

            double[] cutbacks = new double[vertices.Count];
            
            Curve polyCurve = poly as Curve;
            List<InlineComponent> inlineComps = new List<InlineComponent>();
            
            if (polyCurve != null)
            {
                foreach (ObjectId id in btr)
                {
                    if (id.ObjectClass.DxfName == "INSERT")
                    {
                        BlockReference bref = tr.GetObject(id, OpenMode.ForRead) as BlockReference;
                        if (bref != null)
                        {
                            string bLineNumber = "";
                            string bCompType = "";
                            string bTargetDN = "";
                            string bGrandDN = "";
                            
                            foreach (ObjectId attId in bref.AttributeCollection)
                            {
                                AttributeReference attRef = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                                if (attRef != null)
                                {
                                    string tag = attRef.Tag.ToUpper();
                                    if (tag == "NO LIGNE") bLineNumber = attRef.TextString;
                                    else if (tag == "DESIGNATION") bCompType = attRef.TextString;
                                    else if (tag == "PETIT DN") bTargetDN = attRef.TextString;
                                    else if (tag == "GRAND DN") bGrandDN = attRef.TextString;
                                }
                            }
                            
                            if (bLineNumber == lineNumber && bCompType != "COUDE" && !bCompType.Contains("ELBOW"))
                            {
                                List<Point3d> ports = new List<Point3d>();
                                BlockTableRecord btrBlock = tr.GetObject(bref.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                                foreach (ObjectId entId in btrBlock)
                                {
                                    Entity bEnt = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                                    if (bEnt is DBPoint pt)
                                    {
                                        ports.Add(pt.Position.TransformBy(bref.BlockTransform));
                                    }
                                }
                                
                                if (ports.Count >= 2)
                                {
                                    try
                                    {
                                        Point3d closest1 = polyCurve.GetClosestPointTo(ports[0], false);
                                        Point3d closest2 = polyCurve.GetClosestPointTo(ports[1], false);
                                        
                                        if (closest1.DistanceTo(ports[0]) < 1.0 && closest2.DistanceTo(ports[1]) < 1.0)
                                        {
                                            double param1 = polyCurve.GetParameterAtPoint(closest1);
                                            double param2 = polyCurve.GetParameterAtPoint(closest2);
                                            
                                            InlineComponent ic = new InlineComponent
                                            {
                                                Param1 = Math.Min(param1, param2),
                                                Param2 = Math.Max(param1, param2),
                                                Pt1 = param1 < param2 ? closest1 : closest2,
                                                Pt2 = param1 < param2 ? closest2 : closest1,
                                                CompType = bCompType
                                            };
                                            
                                            if (bCompType.Contains("REDUCER") || bCompType.Contains("RED_CONC") || bCompType.Contains("RED_EXC"))
                                            {
                                                ic.TargetDN = param1 < param2 ? bTargetDN : bGrandDN;
                                                
                                                if (LogiK3D.UI.MainPaletteControl.AvailableDiameters != null && 
                                                    LogiK3D.UI.MainPaletteControl.AvailableDiameters.ContainsKey(ic.TargetDN))
                                                {
                                                    ic.TargetOD = LogiK3D.UI.MainPaletteControl.AvailableDiameters[ic.TargetDN];
                                                }
                                                else
                                                {
                                                    string dnStr = ic.TargetDN.Replace("DN", "");
                                                    if (double.TryParse(dnStr, out double dnVal))
                                                    {
                                                        ic.TargetOD = dnVal > 50 ? dnVal + 10 : dnVal + 5;
                                                    }
                                                }
                                            }
                                            
                                            inlineComps.Add(ic);
                                        }
                                    }
                                    catch { }
                                }
                            }
                        }
                    }
                }
                inlineComps.Sort((a, b) => a.Param1.CompareTo(b.Param1));
            }

            // Coudes
              PipingGenerator generator = new PipingGenerator();
              for (int i = 1; i < vertices.Count - 1; i++)
              {
                  // Calculer l'angle et le cutback
                  Vector3d vIn = (vertices[i] - vertices[i - 1]).GetNormal();
                  Vector3d vOut = (vertices[i + 1] - vertices[i]).GetNormal();
                  double angleRad = vIn.GetAngleTo(vOut);
                  
                  if (angleRad < 0.01 || angleRad > Math.PI - 0.01)
                  {
                      cutbacks[i] = 0;
                      continue;
                  }

                  double angleDeg = angleRad * 180.0 / Math.PI;
                  double cutback = elbowRadiusLR * Math.Tan(angleRad / 2.0);
                  cutbacks[i] = cutback;

                  // Créer le bloc coude
                  string blockName = $"ELBOW_{Math.Round(angleDeg)}_{dn}_{elbowRadiusLR}";
                  ObjectId elbowBlockId = generator.GetOrCreateElbow(od, elbowRadiusLR, angleDeg, blockName);

                  if (elbowBlockId != ObjectId.Null)
                  {
                      // Le point d'insertion du bloc est le début du coude (après le cutback)
                      Point3d insertPt = vertices[i] - vIn * cutback;
                      BlockReference bref = new BlockReference(insertPt, elbowBlockId);
                      
                      // Aligner le bloc pour que son axe X corresponde à vIn
                      Vector3d xAxis = Vector3d.XAxis;
                      double angle = xAxis.GetAngleTo(vIn);
                      Vector3d axisOfRotation = xAxis.CrossProduct(vIn);

                      if (!axisOfRotation.IsZeroLength())
                      {
                          bref.TransformBy(Matrix3d.Rotation(angle, axisOfRotation, insertPt));
                      }
                      else if (vIn.X < 0)
                      {
                          bref.TransformBy(Matrix3d.Rotation(Math.PI, Vector3d.ZAxis, insertPt));
                      }
                      
                      // Aligner le plan du coude avec vOut
                      Vector3d currentYAxis = bref.BlockTransform.CoordinateSystem3d.Yaxis;
                      Vector3d targetYAxis = (vOut - vIn * vOut.DotProduct(vIn)).GetNormal();
                      
                      if (!targetYAxis.IsZeroLength())
                      {
                          Plane plane = new Plane(Point3d.Origin, vIn);
                          double currentAngle = currentYAxis.AngleOnPlane(plane);
                          double targetAngle = targetYAxis.AngleOnPlane(plane);
                          bref.TransformBy(Matrix3d.Rotation(targetAngle - currentAngle, vIn, insertPt));
                      }

                      // Renseigner les attributs
                      var attributes = new System.Collections.Generic.Dictionary<string, string>
                      {
                          { "POSITION", "" },
                          { "DESIGNATION", "COUDE" },
                          { "DN", dn },
                          { "PN", "PN16" },
                          { "ID", $"KOH-{dn}-ELBOW" },
                          { "NO LIGNE", lineNumber }
                      };

                      btr.AppendEntity(bref);
                      tr.AddNewlyCreatedDBObject(bref, true);
                      newSolidHandles.Add(bref.Handle);

                      BlockTableRecord btrBlock = (BlockTableRecord)tr.GetObject(elbowBlockId, OpenMode.ForRead);
                      foreach (ObjectId entId in btrBlock)
                      {
                          Entity ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                          if (ent is AttributeDefinition attrDef)
                          {
                              AttributeReference attrRef = new AttributeReference();
                              attrRef.SetAttributeFromBlock(attrDef, bref.BlockTransform);

                              if (attributes.ContainsKey(attrDef.Tag.ToUpper()))
                              {
                                  attrRef.TextString = attributes[attrDef.Tag.ToUpper()];
                              }

                              bref.AttributeCollection.AppendAttribute(attrRef);
                              tr.AddNewlyCreatedDBObject(attrRef, true);
                          }
                      }
                  }
              }

            double currentPipeRadius = pipeRadius;
            string currentPipeDN = dn;
            double currentPipeDNValue = dnValue;

            // Tubes droits
            for (int i = 0; i < vertices.Count - 1; i++)
            {
                Vector3d direction = (vertices[i + 1] - vertices[i]).GetNormal();
                Point3d startPt = vertices[i] + direction * cutbacks[i];
                Point3d endPt = vertices[i + 1] - direction * cutbacks[i + 1];

                double paramStart = polyCurve != null ? polyCurve.GetParameterAtPoint(polyCurve.GetClosestPointTo(startPt, false)) : i;
                double paramEnd = polyCurve != null ? polyCurve.GetParameterAtPoint(polyCurve.GetClosestPointTo(endPt, false)) : i + 1;

                List<InlineComponent> segmentComps = new List<InlineComponent>();
                foreach (var ic in inlineComps)
                {
                    if (ic.Param1 >= paramStart - 0.01 && ic.Param2 <= paramEnd + 0.01)
                    {
                        segmentComps.Add(ic);
                    }
                }

                Point3d currentPt = startPt;

                foreach (var ic in segmentComps)
                {
                    double cutLength = currentPt.DistanceTo(ic.Pt1);
                    if (cutLength > 0.1)
                    {
                        Solid3d pipeSolid = CreatePipeSolid(currentPt, ic.Pt1, currentPipeRadius, thickness);
                        AttachSolidData(pipeSolid, $"KOH-{currentPipeDN}-PIPE", cutLength, lineNumber, "PIPE", currentPipeDNValue, currentPt, ic.Pt1, Point3d.Origin, tr, db);
                        btr.AppendEntity(pipeSolid);
                        tr.AddNewlyCreatedDBObject(pipeSolid, true);
                        newSolidHandles.Add(pipeSolid.Handle);
                    }

                    currentPt = ic.Pt2;

                    if (ic.TargetOD > 0)
                    {
                        currentPipeRadius = ic.TargetOD / 2.0;
                        currentPipeDN = ic.TargetDN;
                        double.TryParse(currentPipeDN.Replace("DN", ""), out currentPipeDNValue);
                    }
                }

                double finalCutLength = currentPt.DistanceTo(endPt);
                if (finalCutLength > 0.1)
                {
                    Solid3d pipeSolid = CreatePipeSolid(currentPt, endPt, currentPipeRadius, thickness);
                    AttachSolidData(pipeSolid, $"KOH-{currentPipeDN}-PIPE", finalCutLength, lineNumber, "PIPE", currentPipeDNValue, currentPt, endPt, Point3d.Origin, tr, db);
                    btr.AppendEntity(pipeSolid);
                    tr.AddNewlyCreatedDBObject(pipeSolid, true);
                    newSolidHandles.Add(pipeSolid.Handle);
                }
            }

            // 3. Sauvegarder les handles des nouveaux solides dans le dictionnaire de la polyligne
            SaveLinkedSolids(poly, newSolidHandles, tr);
        }

        public static void DeleteLinkedSolids(Entity poly, Transaction tr)
        {
            if (poly.ExtensionDictionary.IsNull) return;
            
            DBDictionary extDict = (DBDictionary)tr.GetObject(poly.ExtensionDictionary, OpenMode.ForRead);
            if (extDict.Contains(DictName))
            {
                Xrecord xrec = (Xrecord)tr.GetObject(extDict.GetAt(DictName), OpenMode.ForRead);
                foreach (TypedValue tv in xrec.Data)
                {
                    if (tv.TypeCode == (int)DxfCode.Handle)
                    {
                        Handle h = new Handle(Convert.ToInt64(tv.Value.ToString(), 16));
                        if (poly.Database.TryGetObjectId(h, out ObjectId id))
                        {
                            if (!id.IsErased)
                            {
                                DBObject obj = tr.GetObject(id, OpenMode.ForWrite);
                                obj.Erase();
                            }
                        }
                    }
                }
            }
        }

        private static void SaveLinkedSolids(Entity poly, List<Handle> handles, Transaction tr)
        {
            if (poly.ExtensionDictionary.IsNull)
            {
                poly.UpgradeOpen();
                poly.CreateExtensionDictionary();
            }
            
            DBDictionary extDict = (DBDictionary)tr.GetObject(poly.ExtensionDictionary, OpenMode.ForWrite);
            Xrecord xrec = new Xrecord();
            ResultBuffer rb = new ResultBuffer();
            
            foreach (Handle h in handles)
            {
                rb.Add(new TypedValue((int)DxfCode.Handle, h.Value.ToString("X")));
            }
            xrec.Data = rb;
            
            if (extDict.Contains(DictName))
            {
                extDict.Remove(DictName);
            }
            extDict.SetAt(DictName, xrec);
            tr.AddNewlyCreatedDBObject(xrec, true);
        }

        public static void AttachLineData(Entity ent, string lineNumber, string dn, double od, double thickness, Transaction tr, Database db)
        {
            EnsureRegAppExists(db, tr, LineDataAppName);
            ResultBuffer rb = new ResultBuffer(
                new TypedValue((int)DxfCode.ExtendedDataRegAppName, LineDataAppName),
                new TypedValue((int)DxfCode.ExtendedDataAsciiString, lineNumber),
                new TypedValue((int)DxfCode.ExtendedDataAsciiString, dn),
                new TypedValue((int)DxfCode.ExtendedDataReal, od),
                new TypedValue((int)DxfCode.ExtendedDataReal, thickness)
            );
            ent.XData = rb;
        }

        public static void AttachSolidData(Entity ent, string sapCode, double length, string lineNumber, string compType, double dnValue, Point3d p1, Point3d p2, Point3d p3, Transaction tr, Database db)
        {
            EnsureRegAppExists(db, tr, SolidDataAppName);
            ResultBuffer rb = new ResultBuffer(
                new TypedValue((int)DxfCode.ExtendedDataRegAppName, SolidDataAppName),
                new TypedValue((int)DxfCode.ExtendedDataAsciiString, Guid.NewGuid().ToString()),
                new TypedValue((int)DxfCode.ExtendedDataAsciiString, sapCode),
                new TypedValue((int)DxfCode.ExtendedDataReal, length),
                new TypedValue((int)DxfCode.ExtendedDataAsciiString, lineNumber),
                new TypedValue((int)DxfCode.ExtendedDataAsciiString, compType),
                new TypedValue((int)DxfCode.ExtendedDataReal, dnValue),
                new TypedValue((int)DxfCode.ExtendedDataXCoordinate, p1),
                new TypedValue((int)DxfCode.ExtendedDataXCoordinate, p2),
                new TypedValue((int)DxfCode.ExtendedDataXCoordinate, p3)
            );
            ent.XData = rb;
        }

        public static void EnsureRegAppExists(Database db, Transaction tr, string appName)
        {
            RegAppTable rat = (RegAppTable)tr.GetObject(db.RegAppTableId, OpenMode.ForRead);
            if (!rat.Has(appName))
            {
                rat.UpgradeOpen();
                RegAppTableRecord ratr = new RegAppTableRecord { Name = appName };
                rat.Add(ratr);
                tr.AddNewlyCreatedDBObject(ratr, true);
            }
        }

        public static List<Point3d> GetPolylineVertices(Entity ent, Transaction tr)
        {
            List<Point3d> pts = new List<Point3d>();
            if (ent is Polyline poly)
            {
                for (int i = 0; i < poly.NumberOfVertices; i++)
                    pts.Add(poly.GetPoint3dAt(i));
            }
            else if (ent is Polyline3d poly3d)
            {
                foreach (ObjectId vId in poly3d)
                {
                    PolylineVertex3d v3d = (PolylineVertex3d)tr.GetObject(vId, OpenMode.ForRead);
                    pts.Add(v3d.Position);
                }
            }
            return pts;
        }

        public static Solid3d CreatePipeSolid(Point3d startPt, Point3d endPt, double radius, double thickness)
        {
            double length = startPt.DistanceTo(endPt);
            Vector3d direction = (endPt - startPt).GetNormal();
            Point3d midPt = startPt + (direction * (length / 2.0));

            Solid3d pipe = new Solid3d();
            pipe.CreateFrustum(length, radius, radius, radius);

            if (thickness > 0 && thickness < radius)
            {
                Solid3d innerPipe = new Solid3d();
                innerPipe.CreateFrustum(length + 2.0, radius - thickness, radius - thickness, radius - thickness);
                pipe.BooleanOperation(BooleanOperationType.BoolSubtract, innerPipe);
            }

            Vector3d zAxis = direction;
            Vector3d xAxis = zAxis.IsParallelTo(Vector3d.XAxis) ? Vector3d.YAxis : zAxis.CrossProduct(Vector3d.XAxis).GetNormal();
            Vector3d yAxis = zAxis.CrossProduct(xAxis).GetNormal();

            Matrix3d transform = Matrix3d.AlignCoordinateSystem(
                Point3d.Origin, Vector3d.XAxis, Vector3d.YAxis, Vector3d.ZAxis,
                midPt, xAxis, yAxis, zAxis);

            pipe.TransformBy(transform);
            return pipe;
        }

        public static Solid3d CreateElbowSolid(Point3d p1, Point3d p2, Point3d p3, double pipeRadius, double thickness, double elbowRadius, out double cutback, out Point3d ptA, out Point3d ptB)
        {
            ptA = Point3d.Origin;
            ptB = Point3d.Origin;
            Vector3d vIn = (p2 - p1).GetNormal();
            Vector3d vOut = (p3 - p2).GetNormal();
            
            double angleRad = vIn.GetAngleTo(vOut);
            if (angleRad < 0.01 || angleRad > Math.PI - 0.01) 
            {
                cutback = 0;
                return null;
            }
            
            cutback = elbowRadius * Math.Tan(angleRad / 2.0);
            
            ptA = p2 - vIn * cutback;
            ptB = p2 + vOut * cutback;
            
            Vector3d normal = vIn.CrossProduct(vOut).GetNormal();
            Vector3d dirToCenter = vIn.RotateBy(Math.PI / 2.0, normal);
            Point3d center = ptA + dirToCenter * elbowRadius;
            
            Vector3d vecCA = ptA - center;
            Vector3d vecCB = ptB - center;
            
            Plane plane = new Plane(center, normal);
            double startAngle = vecCA.AngleOnPlane(plane);
            double endAngle = vecCB.AngleOnPlane(plane);
            
            try
            {
                using (Arc arc = new Arc(center, normal, elbowRadius, startAngle, endAngle))
                using (Circle circle = new Circle(ptA, vIn, pipeRadius))
                {
                    DBObjectCollection col = new DBObjectCollection();
                    col.Add(circle);
                    
                    if (thickness > 0 && thickness < pipeRadius)
                    {
                        using (Circle innerCircle = new Circle(ptA, vIn, pipeRadius - thickness))
                        {
                            col.Add(innerCircle);
                            DBObjectCollection regions = Region.CreateFromCurves(col);
                            if (regions.Count >= 2)
                            {
                                Region outerReg = (Region)regions[0];
                                Region innerReg = (Region)regions[1];
                                // Ensure outerReg is actually the larger one
                                if (outerReg.Area < innerReg.Area)
                                {
                                    Region temp = outerReg;
                                    outerReg = innerReg;
                                    innerReg = temp;
                                }
                                outerReg.BooleanOperation(BooleanOperationType.BoolSubtract, innerReg);
                                
                                Solid3d elbow = new Solid3d();
                                SweepOptionsBuilder sob = new SweepOptionsBuilder();
                                sob.Align = SweepOptionsAlignOption.NoAlignment;
                                sob.BasePoint = ptA;
                                
                                elbow.CreateSweptSolid(outerReg, arc, sob.ToSweepOptions());
                                
                                foreach(DBObject obj in regions) obj.Dispose();
                                return elbow;
                            }
                            else
                            {
                                foreach(DBObject obj in regions) obj.Dispose();
                            }
                        }
                    }
                    else
                    {
                        DBObjectCollection regions = Region.CreateFromCurves(col);
                        if (regions.Count > 0)
                        {
                            Region reg = (Region)regions[0];
                            
                            Solid3d elbow = new Solid3d();
                            SweepOptionsBuilder sob = new SweepOptionsBuilder();
                            sob.Align = SweepOptionsAlignOption.NoAlignment;
                            sob.BasePoint = ptA;
                            
                            elbow.CreateSweptSolid(reg, arc, sob.ToSweepOptions());
                            
                            foreach(DBObject obj in regions) obj.Dispose();
                            return elbow;
                        }
                    }
                    return null;
                }
            }
            catch
            {
                return null;
            }
        }
    }
}