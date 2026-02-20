using System;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
using LogiK3D.UI;

namespace LogiK3D
{
    public class PluginInitialization : IExtensionApplication
    {
        // La palette d'outils AutoCAD
        private static PaletteSet _paletteSet = null;
        private static MainPaletteControl _mainControl = null;

        public void Initialize()
        {
            // Ce code s'exécute au chargement de la DLL (netload)
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("\nLogiK 3D chargé avec succès. Tapez LOGIK_PALETTE pour ouvrir l'interface.\n");
        }

        public void Terminate()
        {
            // Nettoyage à la fermeture
            if (_paletteSet != null)
            {
                _paletteSet.Dispose();
                _paletteSet = null;
            }
        }

        [CommandMethod("LOGIK_PALETTE")]
        public void ShowPalette()
        {
            if (_paletteSet == null)
            {
                // Création du PaletteSet
                _paletteSet = new PaletteSet("LogiK 3D - Piping", new Guid("A1B2C3D4-E5F6-7890-1234-567890ABCDEF"));
                
                // Configuration de la palette
                _paletteSet.Style = PaletteSetStyles.ShowPropertiesMenu | 
                                    PaletteSetStyles.ShowAutoHideButton | 
                                    PaletteSetStyles.ShowCloseButton;
                _paletteSet.MinimumSize = new System.Drawing.Size(250, 400);
                _paletteSet.DockEnabled = DockSides.Left | DockSides.Right;

                // Création du contrôle WPF
                _mainControl = new MainPaletteControl();

                // Ajout du contrôle WPF à la palette via un ElementHost
                System.Windows.Forms.Integration.ElementHost host = new System.Windows.Forms.Integration.ElementHost();
                host.AutoSize = true;
                host.Dock = System.Windows.Forms.DockStyle.Fill;
                host.Child = _mainControl;

                _paletteSet.Add("Outils de Tuyauterie", host);
            }

            // Afficher la palette
            _paletteSet.Visible = true;
        }
    }
}