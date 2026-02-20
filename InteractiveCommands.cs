using System;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

namespace LogiK3D.Piping
{
    public class InteractiveCommands
    {
        [CommandMethod("LOGIK_ROUTE_PIPE")]
        public void RoutePipeInteractive()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // 1. Demander le point de départ
            PromptPointOptions ppoStart = new PromptPointOptions("\nSpécifiez le point de départ du tube : ");
            PromptPointResult pprStart = ed.GetPoint(ppoStart);

            if (pprStart.Status != PromptStatus.OK) return;

            Point3d currentPoint = pprStart.Value;

            // 2. Boucle pour tracer les segments
            while (true)
            {
                PromptPointOptions ppoNext = new PromptPointOptions("\nSpécifiez le point suivant (ou Entrée pour terminer) : ");
                ppoNext.UseBasePoint = true;
                ppoNext.BasePoint = currentPoint;
                ppoNext.AllowNone = true; // Permet d'appuyer sur Entrée pour quitter

                PromptPointResult pprNext = ed.GetPoint(ppoNext);

                if (pprNext.Status == PromptStatus.None)
                {
                    // L'utilisateur a appuyé sur Entrée, on arrête
                    break;
                }
                else if (pprNext.Status != PromptStatus.OK)
                {
                    // Annulation (Echap)
                    break;
                }

                Point3d nextPoint = pprNext.Value;

                // 3. Dessiner le segment de tube (Solid3d)
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                    // Création d'un cylindre 3D pour le tube
                    double radius = LogiK3D.UI.MainPaletteControl.CurrentOuterDiameter / 2.0;
                    double thickness = LogiK3D.UI.MainPaletteControl.CurrentThickness;
                    double length = currentPoint.DistanceTo(nextPoint);

                    if (length > 0)
                    {
                        Solid3d pipeSolid = new Solid3d();
                        pipeSolid.CreateFrustum(length, radius, radius, radius);

                        if (thickness > 0 && thickness < radius)
                        {
                            Solid3d innerPipe = new Solid3d();
                            innerPipe.CreateFrustum(length + 2.0, radius - thickness, radius - thickness, radius - thickness);
                            pipeSolid.BooleanOperation(BooleanOperationType.BoolSubtract, innerPipe);
                        }

                        // Positionner et orienter le cylindre
                        Vector3d direction = (nextPoint - currentPoint).GetNormal();
                        Point3d midPoint = currentPoint + (direction * (length / 2.0));

                        // Le cylindre est créé le long de l'axe Z par défaut.
                        // On doit le tourner pour l'aligner avec notre direction.
                        Vector3d zAxis = Vector3d.ZAxis;
                        double angle = zAxis.GetAngleTo(direction);
                        Vector3d axisOfRotation = zAxis.CrossProduct(direction);

                        Matrix3d transform = Matrix3d.Displacement(midPoint.GetAsVector());
                        
                        if (!axisOfRotation.IsZeroLength())
                        {
                            transform = transform * Matrix3d.Rotation(angle, axisOfRotation, Point3d.Origin);
                        }
                        else if (direction.Z < 0) // Cas où la direction est exactement -Z
                        {
                            transform = transform * Matrix3d.Rotation(Math.PI, Vector3d.XAxis, Point3d.Origin);
                        }

                        pipeSolid.TransformBy(transform);
                        pipeSolid.ColorIndex = 3; // Vert

                        // Ajout des XData pour l'export PCF
                        AttachLogiKData(pipeSolid, $"KOH-{LogiK3D.UI.MainPaletteControl.CurrentDN}-PIPE", length, tr, db);

                        btr.AppendEntity(pipeSolid);
                        tr.AddNewlyCreatedDBObject(pipeSolid, true);
                    }

                    tr.Commit();
                }

                // Le point suivant devient le point de départ pour le prochain segment
                currentPoint = nextPoint;
            }
        }

        private void AttachLogiKData(Entity ent, string sapCode, double cutLength, Transaction tr, Database db)
        {
            string appName = "LogiK_Data";
            
            // S'assurer que l'application enregistrée existe
            RegAppTable rat = (RegAppTable)tr.GetObject(db.RegAppTableId, OpenMode.ForRead);
            if (!rat.Has(appName))
            {
                rat.UpgradeOpen();
                RegAppTableRecord ratr = new RegAppTableRecord { Name = appName };
                rat.Add(ratr);
                tr.AddNewlyCreatedDBObject(ratr, true);
            }

            ResultBuffer rb = new ResultBuffer(
                new TypedValue((int)DxfCode.ExtendedDataRegAppName, appName),
                new TypedValue((int)DxfCode.ExtendedDataAsciiString, Guid.NewGuid().ToString()),
                new TypedValue((int)DxfCode.ExtendedDataAsciiString, sapCode),
                new TypedValue((int)DxfCode.ExtendedDataReal, cutLength)
            );
            ent.XData = rb;
        }

        // La commande LOGIK_INSERT_COMP a été déplacée dans SmartTubeCommands.cs
        // pour utiliser le vrai générateur de blocs 3D.

        [CommandMethod("LOGIK_GET_INFO")]
        public void GetInfo()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            PromptEntityOptions peo = new PromptEntityOptions("\nSélectionnez un composant LogiK 3D : ");
            PromptEntityResult per = ed.GetEntity(peo);

            if (per.Status == PromptStatus.OK)
            {
                // TODO: Lire les XData de l'objet
                ed.WriteMessage($"\n[Simulation] Lecture des informations de l'objet {per.ObjectId}");
            }
        }
    }
}