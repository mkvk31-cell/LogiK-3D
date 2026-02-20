using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using LogiK3D.Piping;

namespace LogiK3D.UI
{
    public class LineInfo
    {
        public string LineNumber { get; set; }
        public string Spec { get; set; }
        public string DN { get; set; }
        public ObjectId PolylineId { get; set; }
    }

    public partial class MainPaletteControl : UserControl
    {
        // Instance active pour accès depuis les commandes
        public static MainPaletteControl Instance { get; private set; }

        // Propriétés statiques pour partager les données avec les commandes
        public static double CurrentOuterDiameter { get; private set; } = 114.3;
        public static string CurrentDN { get; private set; } = "DN100";
        public static double CurrentThickness { get; private set; } = 3.2;

        public MainPaletteControl()
        {
            InitializeComponent();
            Instance = this;
            LoadDiameters();
            RefreshLineList();
        }

        public void SetCurrentDN(string dn)
        {
            foreach (ComboBoxItem item in CmbDiameter.Items)
            {
                if (item.Content.ToString().StartsWith(dn))
                {
                    CmbDiameter.SelectedItem = item;
                    break;
                }
            }
        }

        public static Dictionary<string, double> AvailableDiameters { get; private set; }

        private void LoadDiameters()
        {
            // Liste des diamètres ISO/DIN (EN 10220 / ISO 4200 / DIN 2448)
            AvailableDiameters = new Dictionary<string, double>
            {
                {"DN10", 17.2},
                {"DN15", 21.3},
                {"DN20", 26.9},
                {"DN25", 33.7},
                {"DN32", 42.4},
                {"DN40", 48.3},
                {"DN50", 60.3},
                {"DN65", 76.1},
                {"DN80", 88.9},
                {"DN100", 114.3},
                {"DN125", 139.7},
                {"DN150", 168.3},
                {"DN200", 219.1},
                {"DN250", 273.0},
                {"DN300", 323.9},
                {"DN350", 355.6},
                {"DN400", 406.4},
                {"DN450", 457.0},
                {"DN500", 508.0},
                {"DN600", 610.0}
            };

            foreach (var kvp in AvailableDiameters)
            {
                ComboBoxItem item = new ComboBoxItem();
                item.Content = $"{kvp.Key} ({kvp.Value} mm)";
                item.Tag = kvp.Value;
                CmbDiameter.Items.Add(item);
                
                // Sélectionner DN100 par défaut
                if (kvp.Key == "DN100")
                {
                    CmbDiameter.SelectedItem = item;
                }
            }

            LoadSpecs();
        }

        private void LoadSpecs()
        {
            CmbSpec.Items.Clear();
            var specs = LogiK3D.Specs.SpecManager.LoadSpecs();
            foreach (var spec in specs)
            {
                ComboBoxItem item = new ComboBoxItem();
                item.Content = spec.Name;
                item.Tag = spec;
                CmbSpec.Items.Add(item);
            }
            if (CmbSpec.Items.Count > 0)
            {
                CmbSpec.SelectedIndex = 0;
            }
        }

        private void CmbSpec_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Logique pour mettre à jour les diamètres/épaisseurs disponibles en fonction de la spec
        }

        private void BtnManageSpecs_Click(object sender, RoutedEventArgs e)
        {
            SpecEditorWindow editor = new SpecEditorWindow();
            if (editor.ShowDialog() == true)
            {
                LoadSpecs(); // Recharger les specs si elles ont été modifiées
            }
        }

        private void CmbDiameter_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CmbDiameter.SelectedItem is ComboBoxItem item)
            {
                CurrentOuterDiameter = (double)item.Tag;
                CurrentDN = item.Content.ToString().Split(' ')[0];
            }
        }

        private void CmbThickness_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CmbThickness != null && CmbThickness.SelectedItem is ComboBoxItem item)
            {
                if (double.TryParse(item.Tag.ToString(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double thickness))
                {
                    CurrentThickness = thickness;
                }
            }
        }

        private void BtnRoutePipe_Click(object sender, RoutedEventArgs e)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc != null)
            {
                // On utilise notre propre commande de routage interactif
                // qui ne nécessite pas d'être dans un projet Plant 3D
                doc.SendStringToExecute("LOGIK_ROUTE_PIPE ", true, false, false);
            }
        }

        private void BtnConvertPolyline_Click(object sender, RoutedEventArgs e)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc != null)
            {
                doc.SendStringToExecute("LOGIK_PIPE ", true, false, false);
            }
        }

        private void BtnUpdateLine_Click(object sender, RoutedEventArgs e)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc != null)
            {
                doc.SendStringToExecute("LOGIK_UPDATE_LINE ", true, false, false);
            }
        }

        private void BtnRefreshLines_Click(object sender, RoutedEventArgs e)
        {
            RefreshLineList();
        }

        public void RefreshLineList()
        {
            try
            {
                Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                if (doc == null) return;

                Database db = doc.Database;
                List<LineInfo> lines = new List<LineInfo>();

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);
                    foreach (ObjectId id in btr)
                    {
                        if (id.ObjectClass.DxfName == "POLYLINE" || id.ObjectClass.DxfName == "LWPOLYLINE" || id.ObjectClass.DxfName == "POLYLINE3D")
                        {
                            Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                            if (ent != null)
                            {
                                ResultBuffer rb = ent.GetXDataForApplication(PipeManager.LineDataAppName);
                                if (rb != null)
                                {
                                    TypedValue[] values = rb.AsArray();
                                    if (values.Length >= 4)
                                    {
                                        lines.Add(new LineInfo
                                        {
                                            LineNumber = values[1].Value.ToString(),
                                            Spec = values[2].Value.ToString(),
                                            DN = values[3].Value.ToString(),
                                            PolylineId = id
                                        });
                                    }
                                }
                            }
                        }
                    }
                    tr.Commit();
                }

                DgLines.ItemsSource = lines;
            }
            catch { }
        }

        private void DgLines_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (DgLines.SelectedItem is LineInfo selectedLine)
            {
                // Try to select the spec
                foreach (var item in CmbSpec.Items)
                {
                    if (item.ToString() == selectedLine.Spec)
                    {
                        CmbSpec.SelectedItem = item;
                        break;
                    }
                }

                // Try to select the DN
                SetCurrentDN(selectedLine.DN);
            }
        }

        private void BtnSelectLine_Click(object sender, RoutedEventArgs e)
        {
            if (DgLines.SelectedItem is LineInfo selectedLine)
            {
                Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                if (doc != null)
                {
                    Editor ed = doc.Editor;
                    ObjectId[] ids = new ObjectId[] { selectedLine.PolylineId };
                    ed.SetImpliedSelection(ids);
                }
            }
        }

        private void BtnInsertComponent_Click(object sender, RoutedEventArgs e)
        {
            Button btn = sender as Button;
            if (btn != null && btn.Tag != null)
            {
                string componentType = btn.Tag.ToString();
                Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                if (doc != null)
                {
                    doc.SendStringToExecute($"LOGIK_INSERT_COMP {componentType} ", true, false, false);
                }
            }
        }

        private void BtnGetInfo_Click(object sender, RoutedEventArgs e)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc != null)
            {
                doc.SendStringToExecute("LOGIK_GET_INFO ", true, false, false);
            }
        }

        private void BtnExportPCF_Click(object sender, RoutedEventArgs e)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc != null)
            {
                // Lance la commande d'export PCF
                doc.SendStringToExecute("LOGIK_EXPORT_PCF ", true, false, false);
            }
        }
    }
}






