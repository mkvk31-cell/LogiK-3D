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
            Point3dCollection points = new Point3dCollection();
            points.Add(currentPoint);

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
                points.Add(nextPoint);
                currentPoint = nextPoint;
            }

            if (points.Count > 1)
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                    Polyline3d poly3d = new Polyline3d();
                    btr.AppendEntity(poly3d);
                    tr.AddNewlyCreatedDBObject(poly3d, true);

                    foreach (Point3d pt in points)
                    {
                        PolylineVertex3d vertex = new PolylineVertex3d(pt);
                        poly3d.AppendVertex(vertex);
                        tr.AddNewlyCreatedDBObject(vertex, true);
                    }

                    // Attacher les données de ligne à la polyligne
                    string lineNumber = "LIGNE_INCONNUE";
                    string currentDN = LogiK3D.UI.MainPaletteControl.CurrentDN;
                    double currentOD = LogiK3D.UI.MainPaletteControl.CurrentOuterDiameter;
                    double currentThickness = LogiK3D.UI.MainPaletteControl.CurrentThickness;

                    LogiK3D.Piping.PipeManager.AttachLineData(poly3d, lineNumber, currentDN, currentOD, currentThickness, tr, db);

                    // Générer la tuyauterie
                    LogiK3D.Piping.PipeManager.GeneratePiping(poly3d, lineNumber, currentDN, currentOD, currentThickness, tr);

                    tr.Commit();
                }
            }
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