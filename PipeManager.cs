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
            
            // Coudes
            for (int i = 1; i < vertices.Count - 1; i++)
            {
                Solid3d elbowSolid = CreateElbowSolid(vertices[i - 1], vertices[i], vertices[i + 1], pipeRadius, thickness, elbowRadiusLR, out double cutback, out Point3d ptA, out Point3d ptB);
                if (elbowSolid != null)
                {
                    cutbacks[i] = cutback;
                    elbowSolid.ColorIndex = 4; // Cyan
                    AttachSolidData(elbowSolid, $"KOH-{dn}-ELBOW", elbowRadiusLR, lineNumber, "ELBOW", dnValue, ptA, ptB, vertices[i], tr, db);
                    btr.AppendEntity(elbowSolid);
                    tr.AddNewlyCreatedDBObject(elbowSolid, true);
                    newSolidHandles.Add(elbowSolid.Handle);
                }
            }

            // Tubes droits
            for (int i = 0; i < vertices.Count - 1; i++)
            {
                Vector3d direction = (vertices[i + 1] - vertices[i]).GetNormal();
                Point3d startPt = vertices[i] + direction * cutbacks[i];
                Point3d endPt = vertices[i + 1] - direction * cutbacks[i + 1];

                double cutLength = startPt.DistanceTo(endPt);
                if (cutLength > 0)
                {
                    Solid3d pipeSolid = CreatePipeSolid(startPt, endPt, pipeRadius, thickness);
                    AttachSolidData(pipeSolid, $"KOH-{dn}-PIPE", cutLength, lineNumber, "PIPE", dnValue, startPt, endPt, Point3d.Origin, tr, db);
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