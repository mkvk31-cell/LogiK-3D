using System;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

namespace LogiK3D.Piping
{
    public class PipingGenerator
    {
        /// <summary>
        /// Génère ou récupère un bloc 3D de coude (Elbow) dans la base de données AutoCAD.
        /// </summary>
        /// <param name="diameter">Diamètre extérieur du coude.</param>
        /// <param name="radius">Rayon de courbure du coude (à l'axe).</param>
        /// <param name="angle">Angle du coude en degrés (ex: 90).</param>
        /// <param name="blockName">Nom unique du bloc (ex: ELBOW_90_DN100_R152).</param>
        /// <returns>L'ObjectId du BlockTableRecord créé ou existant.</returns>
        public ObjectId GetOrCreateElbow(double diameter, double radius, double angle, string blockName)
        {
            // Récupération du document actif et de sa base de données
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            ObjectId blockId = ObjectId.Null;

            // Utilisation d'une transaction pour garantir l'intégrité de la base de données
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // 1. Check de la BlockTable
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                // Si le bloc existe déjà, on retourne son ObjectId
                if (bt.Has(blockName))
                {
                    blockId = bt[blockName];
                    tr.Commit();
                    return blockId;
                }

                // 2. Création du "Jig" ou Géométrie Temporaire
                // Le bloc n'existe pas, on doit le créer. On passe la BlockTable en mode écriture.
                bt.UpgradeOpen();

                // Création du nouveau BlockTableRecord
                BlockTableRecord btr = new BlockTableRecord();
                btr.Name = blockName;
                // Le bloc est créé à l'origine (0,0,0) par défaut

                // Conversion de l'angle en radians pour les calculs géométriques
                double angleRad = angle * (Math.PI / 180.0);

                // Création des objets temporaires pour le Sweep (Balayage)
                using (Circle profile = new Circle())
                using (Arc path = new Arc())
                {
                    // Définition du chemin (Arc)
                    path.Center = Point3d.Origin;
                    path.Normal = Vector3d.ZAxis; // L'arc est dessiné dans le plan XY
                    path.Radius = radius;
                    path.StartAngle = 0.0;
                    path.EndAngle = angleRad;

                    // Définition du profil (Cercle)
                    // Créé à l'origine avec normale Z, puis déplacé et orienté
                    profile.Center = Point3d.Origin;
                    profile.Normal = Vector3d.ZAxis;
                    profile.Radius = diameter / 2.0;
                    
                    // Déplacement au point de départ de l'arc et orientation (normale Y)
                    profile.TransformBy(Matrix3d.Displacement(new Vector3d(radius, 0, 0)));
                    profile.TransformBy(Matrix3d.Rotation(Math.PI / 2.0, Vector3d.XAxis, new Point3d(radius, 0, 0)));

                    // Création du solide 3D
                    Solid3d solidElbow = new Solid3d();
                    try
                    {
                        // Options de balayage par défaut
                        SweepOptionsBuilder sweepOpts = new SweepOptionsBuilder();
                        
                        // Génération du solide par balayage du profil le long du chemin
                        solidElbow.CreateSweptSolid(profile, path, sweepOpts.ToSweepOptions());

                        // Ajout du solide au BlockTableRecord
                        btr.AppendEntity(solidElbow);
                    }
                    catch (System.Exception ex)
                    {
                        doc.Editor.WriteMessage($"\nErreur lors de la création du solide balayé (Coude) : {ex.Message}");
                        solidElbow.Dispose();
                    }

                    // 3. Points d'accroche (Ports)
                    // Création des DBPoints aux extrémités pour faciliter l'accrochage (Osnap)
                    DBPoint startPort = new DBPoint(path.StartPoint);
                    DBPoint endPort = new DBPoint(path.EndPoint);
                    
                    // Assignation au calque DEFPOINTS (calque non imprimable standard d'AutoCAD)
                    // Note: Il est recommandé de s'assurer que le calque DEFPOINTS existe dans le dessin,
                    // mais AutoCAD le gère généralement bien si on l'assigne directement.
                    try { startPort.Layer = "DEFPOINTS"; } catch { }
                    try { endPort.Layer = "DEFPOINTS"; } catch { }

                    // Ajout des points au BlockTableRecord
                    btr.AppendEntity(startPort);
                    btr.AppendEntity(endPort);
                } // Les objets temporaires (profile, path) sont automatiquement Dispose() ici

                // 4. Encapsulation dans la Database
                // Ajout du BlockTableRecord à la BlockTable
                blockId = bt.Add(btr);
                
                // Notification à la transaction de l'ajout du nouveau record
                tr.AddNewlyCreatedDBObject(btr, true);

                // Validation de la transaction
                tr.Commit();
            }

            return blockId;
        }

        /// <summary>
        /// Génère ou récupère un bloc 3D de Té (Tee) dans la base de données AutoCAD.
        /// </summary>
        /// <param name="mainDiameter">Diamètre extérieur du tube principal.</param>
        /// <param name="branchDiameter">Diamètre extérieur de la dérivation.</param>
        /// <param name="length">Longueur totale du corps principal.</param>
        /// <param name="branchHeight">Hauteur de la dérivation depuis l'axe central.</param>
        /// <param name="blockName">Nom unique du bloc (ex: TEE_EGAL_DN100).</param>
        /// <returns>L'ObjectId du BlockTableRecord créé ou existant.</returns>
        public ObjectId GetOrCreateTee(double mainDiameter, double branchDiameter, double length, double branchHeight, string blockName)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            ObjectId blockId = ObjectId.Null;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                if (bt.Has(blockName))
                {
                    blockId = bt[blockName];
                    tr.Commit();
                    return blockId;
                }

                bt.UpgradeOpen();
                BlockTableRecord btr = new BlockTableRecord();
                btr.Name = blockName;

                // Création du corps principal (Cylindre aligné sur l'axe X)
                Solid3d mainSolid = new Solid3d();
                try
                {
                    // CreateFrustum crée un cylindre le long de l'axe Z centré en 0,0,0
                    // Paramètres: height, xRadius, yRadius, topXRadius
                    mainSolid.CreateFrustum(length, mainDiameter / 2.0, mainDiameter / 2.0, mainDiameter / 2.0);
                    
                    // Rotation pour l'aligner sur l'axe X
                    mainSolid.TransformBy(Matrix3d.Rotation(Math.PI / 2.0, Vector3d.YAxis, Point3d.Origin));

                    // Création de la branche (Cylindre le long de l'axe Y)
                    Solid3d branchSolid = new Solid3d();
                    branchSolid.CreateFrustum(branchHeight, branchDiameter / 2.0, branchDiameter / 2.0, branchDiameter / 2.0);
                    
                    // Rotation pour l'aligner sur l'axe Y
                    branchSolid.TransformBy(Matrix3d.Rotation(Math.PI / 2.0, Vector3d.XAxis, Point3d.Origin)); 
                    
                    // Décalage pour que la base commence à l'axe (0,0,0) et aille jusqu'à branchHeight
                    branchSolid.TransformBy(Matrix3d.Displacement(new Vector3d(0, branchHeight / 2.0, 0)));

                    // Union booléenne : on ajoute la branche au corps principal
                    // ATTENTION: BooleanOperation détruit branchSolid. Ne pas utiliser de 'using' !
                    mainSolid.BooleanOperation(BooleanOperationType.BoolUnite, branchSolid);

                    btr.AppendEntity(mainSolid);
                }
                catch (System.Exception ex)
                {
                    doc.Editor.WriteMessage($"\nErreur lors de la création du solide (Té) : {ex.Message}");
                    mainSolid.Dispose();
                }

                // Points d'accroche (Ports)
                DBPoint port1 = new DBPoint(new Point3d(-length / 2.0, 0, 0));
                DBPoint port2 = new DBPoint(new Point3d(length / 2.0, 0, 0));
                DBPoint port3 = new DBPoint(new Point3d(0, branchHeight, 0));
                
                try { port1.Layer = "DEFPOINTS"; } catch { }
                try { port2.Layer = "DEFPOINTS"; } catch { }
                try { port3.Layer = "DEFPOINTS"; } catch { }
                
                btr.AppendEntity(port1);
                btr.AppendEntity(port2);
                btr.AppendEntity(port3);

                blockId = bt.Add(btr);
                tr.AddNewlyCreatedDBObject(btr, true);
                tr.Commit();
            }
            return blockId;
        }

        /// <summary>
        /// Génère ou récupère un bloc 3D de Réduction (Reducer) concentrique.
        /// </summary>
        /// <param name="largeDiameter">Grand diamètre extérieur.</param>
        /// <param name="smallDiameter">Petit diamètre extérieur.</param>
        /// <param name="length">Longueur totale de la réduction.</param>
        /// <param name="blockName">Nom unique du bloc (ex: REDUCER_DN100_DN80).</param>
        /// <returns>L'ObjectId du BlockTableRecord créé ou existant.</returns>
        public ObjectId GetOrCreateReducer(double largeDiameter, double smallDiameter, double length, string blockName)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            ObjectId blockId = ObjectId.Null;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                if (bt.Has(blockName))
                {
                    blockId = bt[blockName];
                    tr.Commit();
                    return blockId;
                }

                bt.UpgradeOpen();
                BlockTableRecord btr = new BlockTableRecord();
                btr.Name = blockName;

                Solid3d reducerSolid = new Solid3d();
                try
                {
                    // Tronc de cône le long de l'axe Z
                    // Paramètres: height, xRadius, yRadius, topXRadius
                    // Note: CreateFrustum in AutoCAD .NET takes (height, xRadius, yRadius, topXRadius)
                    // If topXRadius is different from xRadius, it creates a cone.
                    reducerSolid.CreateFrustum(length, largeDiameter / 2.0, largeDiameter / 2.0, smallDiameter / 2.0);
                    
                    // Rotation pour l'aligner sur l'axe X
                    reducerSolid.TransformBy(Matrix3d.Rotation(Math.PI / 2.0, Vector3d.YAxis, Point3d.Origin));

                    btr.AppendEntity(reducerSolid);
                }
                catch (System.Exception ex)
                {
                    doc.Editor.WriteMessage($"\nErreur lors de la création du solide (Réduction) : {ex.Message}");
                    reducerSolid.Dispose();
                }

                // Points d'accroche (Ports)
                DBPoint port1 = new DBPoint(new Point3d(-length / 2.0, 0, 0)); // Côté large
                DBPoint port2 = new DBPoint(new Point3d(length / 2.0, 0, 0));  // Côté réduit
                
                try { port1.Layer = "DEFPOINTS"; } catch { }
                try { port2.Layer = "DEFPOINTS"; } catch { }
                
                btr.AppendEntity(port1);
                btr.AppendEntity(port2);

                blockId = bt.Add(btr);
                tr.AddNewlyCreatedDBObject(btr, true);
                tr.Commit();
            }
            return blockId;
        }

        /// <summary>
        /// Génère ou récupère un bloc 3D de Bride (Flange).
        /// </summary>
        /// <param name="pipeDiameter">Diamètre extérieur du tube de raccordement.</param>
        /// <param name="flangeDiameter">Diamètre extérieur de la face de la bride.</param>
        /// <param name="thickness">Épaisseur de la face de la bride.</param>
        /// <param name="length">Longueur totale (face + collet).</param>
        /// <param name="blockName">Nom unique du bloc (ex: FLANGE_DN100_PN16).</param>
        /// <returns>L'ObjectId du BlockTableRecord créé ou existant.</returns>
        public ObjectId GetOrCreateFlange(double pipeDiameter, double flangeDiameter, double thickness, double length, string blockName)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            ObjectId blockId = ObjectId.Null;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                if (bt.Has(blockName))
                {
                    blockId = bt[blockName];
                    tr.Commit();
                    return blockId;
                }

                bt.UpgradeOpen();
                BlockTableRecord btr = new BlockTableRecord();
                btr.Name = blockName;

                Solid3d flangeFace = new Solid3d();
                try
                {
                    // Face de la bride (disque large et fin)
                    // Paramètres: height, xRadius, yRadius, topXRadius
                    flangeFace.CreateFrustum(thickness, flangeDiameter / 2.0, flangeDiameter / 2.0, flangeDiameter / 2.0);
                    
                    // Col de la bride (cylindre au diamètre du tube)
                    double neckLength = length - thickness;
                    if (neckLength > 0)
                    {
                        Solid3d neck = new Solid3d();
                        neck.CreateFrustum(neckLength, pipeDiameter / 2.0, pipeDiameter / 2.0, pipeDiameter / 2.0);
                        // Décalage du col pour qu'il soit collé à la face
                        neck.TransformBy(Matrix3d.Displacement(new Vector3d(0, 0, (thickness + neckLength) / 2.0)));
                        
                        // Union
                        // ATTENTION: BooleanOperation détruit neck. Ne pas utiliser de 'using' !
                        flangeFace.BooleanOperation(BooleanOperationType.BoolUnite, neck);
                    }

                    // Rotation pour aligner sur l'axe X
                    flangeFace.TransformBy(Matrix3d.Rotation(Math.PI / 2.0, Vector3d.YAxis, Point3d.Origin));

                    btr.AppendEntity(flangeFace);
                }
                catch (System.Exception ex)
                {
                    doc.Editor.WriteMessage($"\nErreur lors de la création du solide (Bride) : {ex.Message}");
                    flangeFace.Dispose();
                }

                // Points d'accroche (Ports)
                // Port 1 : Face de la bride (connexion avec une autre bride)
                // Port 2 : Bout du col (connexion avec le tube)
                DBPoint port1 = new DBPoint(new Point3d(-thickness / 2.0, 0, 0));
                DBPoint port2 = new DBPoint(new Point3d(-thickness / 2.0 + length, 0, 0));
                
                try { port1.Layer = "DEFPOINTS"; } catch { }
                try { port2.Layer = "DEFPOINTS"; } catch { }
                
                btr.AppendEntity(port1);
                btr.AppendEntity(port2);

                blockId = bt.Add(btr);
                tr.AddNewlyCreatedDBObject(btr, true);
                tr.Commit();
            }
            return blockId;
        }

        /// <summary>
        /// Insère une référence de bloc dans le ModelSpace.
        /// </summary>
        /// <param name="blockId">L'ObjectId du bloc à insérer.</param>
        /// <param name="position">La position d'insertion.</param>
        /// <returns>L'ObjectId de la BlockReference insérée.</returns>
        public ObjectId InsertBlockReference(ObjectId blockId, Point3d position)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            ObjectId brefId = ObjectId.Null;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                BlockReference bref = new BlockReference(position, blockId);
                brefId = modelSpace.AppendEntity(bref);
                tr.AddNewlyCreatedDBObject(bref, true);

                tr.Commit();
            }

            return brefId;
        }
    }
}