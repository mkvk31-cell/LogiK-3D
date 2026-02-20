using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using LogiK3D.Specs;

namespace LogiK3D.UI
{
    public partial class SpecEditorWindow : Window
    {
        private ObservableCollection<PipingSpec> _specs;
        private PipingSpec _currentSpec;
        private ObservableCollection<SpecComponent> _currentComponents;

        public SpecEditorWindow()
        {
            InitializeComponent();
            LoadData();
        }

        private void LoadData()
        {
            var loadedSpecs = SpecManager.LoadSpecs();
            _specs = new ObservableCollection<PipingSpec>(loadedSpecs);
            LstSpecs.ItemsSource = _specs;
            if (_specs.Count > 0)
            {
                LstSpecs.SelectedIndex = 0;
            }
        }

        private void LstSpecs_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (LstSpecs.SelectedItem is PipingSpec spec)
            {
                _currentSpec = spec;
                TxtSpecName.Text = spec.Name;
                TxtSpecDesc.Text = spec.Description;
                TxtSpecMat.Text = spec.Material;

                _currentComponents = new ObservableCollection<SpecComponent>(spec.Components);
                GridComponents.ItemsSource = _currentComponents;
            }
            else
            {
                _currentSpec = null;
                TxtSpecName.Text = "";
                TxtSpecDesc.Text = "";
                TxtSpecMat.Text = "";
                GridComponents.ItemsSource = null;
            }
        }

        private void TxtSpecField_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_currentSpec != null)
            {
                _currentSpec.Name = TxtSpecName.Text;
                _currentSpec.Description = TxtSpecDesc.Text;
                _currentSpec.Material = TxtSpecMat.Text;
                LstSpecs.Items.Refresh();
            }
        }

        private void BtnAddSpec_Click(object sender, RoutedEventArgs e)
        {
            var newSpec = new PipingSpec { Name = "Nouvelle Spec", Description = "", Material = "" };
            _specs.Add(newSpec);
            LstSpecs.SelectedItem = newSpec;
        }

        private void BtnImportPdf_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.DefaultExt = ".pdf";
            dlg.Filter = "Fichiers PDF (*.pdf)|*.pdf";

            if (dlg.ShowDialog() == true)
            {
                string filename = dlg.FileName;
                List<PipingSpec> importedSpecs = PdfSpecParser.ParsePdf(filename);
                
                if (importedSpecs != null && importedSpecs.Count > 0)
                {
                    foreach (var spec in importedSpecs)
                    {
                        _specs.Add(spec);
                    }
                    LstSpecs.SelectedItem = importedSpecs.Last();
                    MessageBox.Show($"{importedSpecs.Count} spécification(s) importée(s) avec succès.\n\nNote: Le parsing PDF est basique et doit être adapté au format exact de vos documents.", "Import PDF", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
        }

        private void BtnDeleteSpec_Click(object sender, RoutedEventArgs e)
        {
            if (LstSpecs.SelectedItem is PipingSpec spec)
            {
                _specs.Remove(spec);
            }
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            // Mettre à jour les composants de la spec courante avant de sauvegarder
            if (_currentSpec != null && _currentComponents != null)
            {
                _currentSpec.Components = _currentComponents.ToList();
            }

            // Sauvegarder toutes les specs
            SpecManager.SaveSpecs(_specs.ToList());
            MessageBox.Show("Spécifications enregistrées avec succès.", "LogiK 3D", MessageBoxButton.OK, MessageBoxImage.Information);
            this.DialogResult = true;
            this.Close();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }
    }
}
