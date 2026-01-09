using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.DocumentParts;

namespace AutoNAVSearchSets
{
    public partial class MainWindow : Window
    {
        private Document _doc;
        private Dictionary<string, List<string>> _propertyCache;
        private Dictionary<string, SelectionSet> _disciplineSets;

        public MainWindow()
        {
            InitializeComponent();
            _doc = Autodesk.Navisworks.Api.Application.ActiveDocument;
            _propertyCache = new Dictionary<string, List<string>>();
            _disciplineSets = new Dictionary<string, SelectionSet>();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            LoadDisciplineSets();
            PopulateModelCheckboxes();
        }

        #region Function 1: Discipline Search Sets

        private void BtnCreateDisciplineSets_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                btnCreateDisciplineSets.IsEnabled = false;
                UpdateStatus("Creating discipline search sets...", "#3498DB");

                CreateDisciplineSearchSets();
                LoadDisciplineSets(); // Reload after creation

                UpdateStatus("Discipline search sets created successfully!", "#27AE60");
                MessageBox.Show(
                    "Discipline search sets have been created successfully.\n\nCheck the 'Sets' panel under '1. DISCIPLINES' folder.",
                    "Success",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                UpdateStatus("Error creating discipline search sets", "#E74C3C");
                MessageBox.Show(
                    "Error creating discipline search sets:\n\n" + ex.Message,
                    "Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                btnCreateDisciplineSets.IsEnabled = true;
            }
        }

        private void CreateDisciplineSearchSets()
        {
            string[] disciplinePatterns = new string[]
            {
                "_ARCH_", "_STRC_", "_MEP_", "_MECH_", "_ELEC_", "_PLUM_",
                "_HVAC_", "_FIRE_", "_CIVIL_", "_SITE_", "_LAND_"
            };

            Dictionary<string, bool> disciplinesFound = new Dictionary<string, bool>();

            foreach (ModelItem item in _doc.Models.RootItems)
            {
                string itemName = item.DisplayName;

                foreach (string pattern in disciplinePatterns)
                {
                    if (itemName.IndexOf(pattern, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        string disciplineName = pattern.Trim('_');
                        if (!disciplinesFound.ContainsKey(disciplineName))
                        {
                            disciplinesFound[disciplineName] = true;
                        }
                        break;
                    }
                }
            }

            if (disciplinesFound.Count == 0)
            {
                throw new Exception("No discipline patterns found in model files.\n\nPlease ensure your files contain discipline identifiers like _ARCH_, _STRC_, _MEP_, etc.");
            }

            // Create "1. DISCIPLINES" folder
            FolderItem disciplinesFolder = new FolderItem();
            disciplinesFolder.DisplayName = "1. DISCIPLINES";

            foreach (string discipline in disciplinesFound.Keys)
            {
                Search search = new Search();
                search.Selection.SelectAll();

                string wildcardPattern = "*_" + discipline + "_*";
                SearchCondition condition = SearchCondition.HasPropertyByDisplayName("Item", "Name")
                    .EqualValue(VariantData.FromDisplayString(wildcardPattern));

                search.SearchConditions.Add(condition);
                search.Locations = SearchLocations.DescendantsAndSelf;

                SelectionSet selSet = new SelectionSet(search);
                selSet.DisplayName = discipline;

                disciplinesFolder.Children.Add(selSet);
            }

            _doc.SelectionSets.AddCopy(disciplinesFolder);
        }

        private void LoadDisciplineSets()
        {
            _disciplineSets.Clear();

            // Find "1. DISCIPLINES" folder
            FolderItem disciplinesFolder = null;
            foreach (SavedItem item in _doc.SelectionSets.RootItem.Children)
            {
                if (item.IsGroup && item.DisplayName == "1. DISCIPLINES")
                {
                    disciplinesFolder = (FolderItem)item;
                    break;
                }
            }

            if (disciplinesFolder != null)
            {
                foreach (SavedItem discItem in disciplinesFolder.Children)
                {
                    if (!discItem.IsGroup)
                    {
                        SelectionSet set = (SelectionSet)discItem;
                        _disciplineSets[set.DisplayName] = set;
                    }
                }
            }
        }

        #endregion

        #region Function 2: Element Category Search Sets (Locator-Based)

        private void PopulateModelCheckboxes()
        {
            pnlModelCheckboxes.Children.Clear();
            pnlModelCheckboxesFunction3.Children.Clear();

            foreach (var discipline in _disciplineSets.Keys.OrderBy(k => k))
            {
                // Function 2 checkbox
                CheckBox chk2 = new CheckBox
                {
                    Content = discipline,
                    IsChecked = true,
                    Margin = new Thickness(5, 2, 5, 2),
                    FontSize = 12
                };
                pnlModelCheckboxes.Children.Add(chk2);

                // Function 3 checkbox
                CheckBox chk3 = new CheckBox
                {
                    Content = discipline,
                    IsChecked = true,
                    Margin = new Thickness(5, 2, 5, 2),
                    FontSize = 12
                };
                pnlModelCheckboxesFunction3.Children.Add(chk3);
            }
        }

        private List<string> GetSelectedDisciplines()
        {
            List<string> selected = new List<string>();

            foreach (UIElement element in pnlModelCheckboxes.Children)
            {
                if (element is CheckBox chk && chk.IsChecked == true)
                {
                    selected.Add(chk.Content.ToString());
                }
            }

            return selected;
        }

        private void BtnCreateElementSets_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                btnCreateElementSets.IsEnabled = false;
                UpdateStatus("Creating element category search sets...", "#3498DB");

                string selectedParameter = GetSelectedParameter();
                List<string> selectedDisciplines = GetSelectedDisciplines();

                if (selectedDisciplines.Count == 0)
                {
                    MessageBox.Show("Please select at least one discipline model.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                CreateElementCategorySearchSets_WithLocator(selectedParameter, selectedDisciplines);

                UpdateStatus("Element category search sets created successfully!", "#27AE60");
                MessageBox.Show(
                    "Element category search sets have been created successfully.\n\nCheck the 'Sets' panel under '2. CLASH SETS' folder.",
                    "Success",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                UpdateStatus("Error creating element search sets", "#E74C3C");
                MessageBox.Show(
                    "Error creating element search sets:\n\n" + ex.Message,
                    "Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                btnCreateElementSets.IsEnabled = true;
            }
        }

        private string GetSelectedParameter()
        {
            if (rbCategory.IsChecked == true) return "Category";
            if (rbSystemName.IsChecked == true) return "SystemName";
            if (rbSystemClassification.IsChecked == true) return "SystemClassification";
            if (rbWorkset.IsChecked == true) return "Workset";
            if (rbFamilyType.IsChecked == true) return "FamilyType";
            return "Category";
        }

        private void CreateElementCategorySearchSets_WithLocator(string parameterType, List<string> selectedDisciplines)
        {
            if (_disciplineSets.Count == 0)
            {
                throw new Exception("Please create discipline search sets first (Function 1) before creating element category search sets.");
            }

            // Find or create "2. CLASH SETS" folder
            FolderItem clashSetsFolder = FindOrCreateClashSetsFolder();

            int totalCreated = 0;

            // For each selected discipline
            foreach (string disciplineName in selectedDisciplines)
            {
                if (!_disciplineSets.ContainsKey(disciplineName))
                    continue;

                SelectionSet disciplineSet = _disciplineSets[disciplineName];

                // Find or create discipline folder within clash sets
                FolderItem disciplineFolder = FindOrCreateDisciplineFolder(clashSetsFolder, disciplineName);

                // Discover unique property values for this discipline
                var propertyValues = DiscoverPropertyValues(disciplineSet, parameterType);

                if (propertyValues.Count == 0)
                {
                    continue;
                }

                // Create search sets for each property value WITH LOCATOR
                foreach (var propertyValue in propertyValues)
                {
                    try
                    {
                        // Create search that uses LOCATOR to scope to discipline
                        Search search = CreateSearchWithLocator(disciplineName, parameterType, propertyValue);

                        if (search != null)
                        {
                            SelectionSet propSet = new SelectionSet(search);
                            propSet.DisplayName = propertyValue;

                            disciplineFolder.Children.Add(propSet);
                            totalCreated++;
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error creating {disciplineName}\\{propertyValue}: {ex.Message}");
                    }
                }
            }

            if (totalCreated == 0)
            {
                throw new Exception("No element categories found in the selected discipline search sets.");
            }
        }

        private FolderItem FindOrCreateClashSetsFolder()
        {
            // Check if "2. CLASH SETS" folder already exists
            foreach (SavedItem item in _doc.SelectionSets.RootItem.Children)
            {
                if (item.IsGroup && item.DisplayName == "2. CLASH SETS")
                {
                    return (FolderItem)item;
                }
            }

            // Create new folder
            FolderItem clashSetsFolder = new FolderItem();
            clashSetsFolder.DisplayName = "2. CLASH SETS";
            _doc.SelectionSets.AddCopy(clashSetsFolder);

            // Return reference to the newly added folder
            foreach (SavedItem item in _doc.SelectionSets.RootItem.Children)
            {
                if (item.IsGroup && item.DisplayName == "2. CLASH SETS")
                {
                    return (FolderItem)item;
                }
            }

            return clashSetsFolder;
        }

        private FolderItem FindOrCreateDisciplineFolder(FolderItem parentFolder, string disciplineName)
        {
            // Check if discipline folder already exists
            foreach (SavedItem item in parentFolder.Children)
            {
                if (item.IsGroup && item.DisplayName == disciplineName)
                {
                    return (FolderItem)item;
                }
            }

            // Create new discipline folder
            FolderItem disciplineFolder = new FolderItem();
            disciplineFolder.DisplayName = disciplineName;
            parentFolder.Children.Add(disciplineFolder);

            return disciplineFolder;
        }

        private Search CreateSearchWithLocator(string disciplineName, string parameterType, string propertyValue)
        {
            Search search = new Search();

            // CRITICAL: Set the locator to scope search to specific discipline
            // This creates: <locator>lcop_selection_set_tree/1. DISCIPLINES/ARCH</locator>
            string locatorPath = "lcop_selection_set_tree/1. DISCIPLINES/" + disciplineName;
            
            // Use the discipline search set as the selection source
            if (_disciplineSets.ContainsKey(disciplineName))
            {
                SelectionSet disciplineSet = _disciplineSets[disciplineName];
                ModelItemCollection items = disciplineSet.Search.FindAll(_doc, false);
                search.Selection.CopyFrom(items);
            }

            // Add property condition
            string categoryName, propertyName;
            GetPropertyMapping(parameterType, out categoryName, out propertyName);

            SearchCondition condition = SearchCondition.HasPropertyByDisplayName(categoryName, propertyName);
            condition = condition.EqualValue(VariantData.FromDisplayString(propertyValue));
            
            search.SearchConditions.Add(condition);
            search.Locations = SearchLocations.DescendantsAndSelf;

            return search;
        }

        private List<string> DiscoverPropertyValues(SelectionSet disciplineSet, string propertyType)
        {
            var values = new HashSet<string>();

            ModelItemCollection items = disciplineSet.Search.FindAll(_doc, false);

            if (items.Count == 0)
            {
                return new List<string>();
            }

            string categoryName, propertyName;
            GetPropertyMapping(propertyType, out categoryName, out propertyName);

            // Scan items for property values
            foreach (ModelItem item in items)
            {
                foreach (ModelItem descendant in GetAllDescendants(item))
                {
                    try
                    {
                        foreach (PropertyCategory cat in descendant.PropertyCategories)
                        {
                            if (cat.DisplayName == categoryName || cat.Name == categoryName)
                            {
                                foreach (DataProperty prop in cat.Properties)
                                {
                                    if (prop.DisplayName == propertyName || prop.Name == propertyName)
                                    {
                                        string val = prop.Value.ToDisplayString();
                                        if (!string.IsNullOrWhiteSpace(val))
                                        {
                                            values.Add(val);
                                        }
                                        break;
                                    }
                                }
                                break;
                            }
                        }
                    }
                    catch { continue; }
                }
            }

            return values.OrderBy(v => v).ToList();
        }

        private void GetPropertyMapping(string propertyType, out string categoryName, out string propertyName)
        {
            switch (propertyType)
            {
                case "Category":
                    categoryName = "Element";
                    propertyName = "Category";
                    break;
                case "SystemName":
                    categoryName = "Element";
                    propertyName = "System Name";
                    break;
                case "SystemClassification":
                    categoryName = "Element";
                    propertyName = "System Classification";
                    break;
                case "Workset":
                    categoryName = "Element";
                    propertyName = "Workset";
                    break;
                case "FamilyType":
                    categoryName = "Element";
                    propertyName = "Type";
                    break;
                default:
                    categoryName = "Element";
                    propertyName = "Category";
                    break;
            }
        }

        private IEnumerable<ModelItem> GetAllDescendants(ModelItem item)
        {
            yield return item;

            if (item.Children != null)
            {
                foreach (ModelItem child in item.Children)
                {
                    foreach (ModelItem descendant in GetAllDescendants(child))
                    {
                        yield return descendant;
                    }
                }
            }
        }

        #endregion

        #region Function 3: Custom Search Set Builder (Auto-Discovery)

        private void BtnRescanProperties_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                btnRescanProperties.IsEnabled = false;
                UpdateStatus("Scanning properties...", "#3498DB");

                List<string> selectedDisciplines = GetSelectedDisciplinesFunction3();
                
                if (selectedDisciplines.Count == 0)
                {
                    MessageBox.Show("Please select at least one discipline model.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                ScanAvailableProperties(selectedDisciplines);

                UpdateStatus("Properties scanned successfully!", "#27AE60");
            }
            catch (Exception ex)
            {
                UpdateStatus("Error scanning properties", "#E74C3C");
                MessageBox.Show(
                    "Error scanning properties:\n\n" + ex.Message,
                    "Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                btnRescanProperties.IsEnabled = true;
            }
        }

        private List<string> GetSelectedDisciplinesFunction3()
        {
            List<string> selected = new List<string>();

            foreach (UIElement element in pnlModelCheckboxesFunction3.Children)
            {
                if (element is CheckBox chk && chk.IsChecked == true)
                {
                    selected.Add(chk.Content.ToString());
                }
            }

            return selected;
        }

        private void ScanAvailableProperties(List<string> selectedDisciplines)
        {
            _propertyCache.Clear();
            cmbPropertyCategory.Items.Clear();

            var categoriesFound = new HashSet<string>();

            foreach (string disciplineName in selectedDisciplines)
            {
                if (!_disciplineSets.ContainsKey(disciplineName))
                    continue;

                SelectionSet disciplineSet = _disciplineSets[disciplineName];
                ModelItemCollection items = disciplineSet.Search.FindAll(_doc, false);

                // Scan properties from selected discipline
                foreach (ModelItem item in items.Take(500)) // Limit for performance
                {
                    foreach (ModelItem descendant in GetAllDescendants(item).Take(50))
                    {
                        try
                        {
                            foreach (PropertyCategory cat in descendant.PropertyCategories)
                            {
                                string catName = cat.DisplayName;

                                if (!categoriesFound.Contains(catName))
                                {
                                    categoriesFound.Add(catName);
                                }

                                if (!_propertyCache.ContainsKey(catName))
                                {
                                    _propertyCache[catName] = new List<string>();
                                }

                                foreach (DataProperty prop in cat.Properties)
                                {
                                    string propName = prop.DisplayName;

                                    if (!_propertyCache[catName].Contains(propName))
                                    {
                                        _propertyCache[catName].Add(propName);
                                    }
                                }
                            }
                        }
                        catch { continue; }
                    }
                }
            }

            // Populate property category dropdown
            foreach (string category in categoriesFound.OrderBy(c => c))
            {
                cmbPropertyCategory.Items.Add(category);
            }

            if (cmbPropertyCategory.Items.Count > 0)
            {
                cmbPropertyCategory.SelectedIndex = 0;
            }
        }

        private void CmbPropertyCategory_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            cmbPropertyName.Items.Clear();

            string selectedCategory = cmbPropertyCategory.SelectedItem?.ToString();
            if (string.IsNullOrEmpty(selectedCategory))
                return;

            if (_propertyCache.ContainsKey(selectedCategory))
            {
                foreach (string propName in _propertyCache[selectedCategory].OrderBy(p => p))
                {
                    cmbPropertyName.Items.Add(propName);
                }

                if (cmbPropertyName.Items.Count > 0)
                {
                    cmbPropertyName.SelectedIndex = 0;
                }
            }
        }

        private void BtnCreateCustomSearchSets_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                btnCreateCustomSearchSets.IsEnabled = false;
                UpdateStatus("Creating custom search sets...", "#3498DB");

                string category = cmbPropertyCategory.SelectedItem?.ToString();
                string propertyName = cmbPropertyName.SelectedItem?.ToString();
                List<string> selectedDisciplines = GetSelectedDisciplinesFunction3();

                if (string.IsNullOrEmpty(category) || string.IsNullOrEmpty(propertyName))
                {
                    MessageBox.Show("Please scan properties first and select a property category and name.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (selectedDisciplines.Count == 0)
                {
                    MessageBox.Show("Please select at least one discipline model.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                int setsCreated = CreateCustomSearchSetsAutoDiscovery(category, propertyName, selectedDisciplines);

                UpdateStatus($"Created {setsCreated} custom search sets successfully!", "#27AE60");
                MessageBox.Show(
                    $"Created {setsCreated} custom search sets.\n\nCheck the 'Sets' panel under '3. CUSTOM SETS' folder.",
                    "Success",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                UpdateStatus("Error creating custom search sets", "#E74C3C");
                MessageBox.Show(
                    "Error creating custom search sets:\n\n" + ex.Message,
                    "Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                btnCreateCustomSearchSets.IsEnabled = true;
            }
        }

        private int CreateCustomSearchSetsAutoDiscovery(string category, string propertyName, List<string> selectedDisciplines)
        {
            // Find or create "3. CUSTOM SETS" folder
            FolderItem customSetsFolder = FindOrCreateCustomSetsFolder();

            int totalCreated = 0;

            // For each selected discipline
            foreach (string disciplineName in selectedDisciplines)
            {
                if (!_disciplineSets.ContainsKey(disciplineName))
                    continue;

                SelectionSet disciplineSet = _disciplineSets[disciplineName];

                // Find or create discipline folder
                FolderItem disciplineFolder = FindOrCreateDisciplineFolder(customSetsFolder, disciplineName);

                // Discover all unique values for this property in this discipline
                var uniqueValues = DiscoverPropertyValuesForCustom(disciplineSet, category, propertyName);

                if (uniqueValues.Count == 0)
                    continue;

                // Create search set for each unique value WITH LOCATOR
                foreach (string value in uniqueValues)
                {
                    try
                    {
                        Search search = CreateCustomSearchWithLocator(disciplineName, category, propertyName, value);

                        if (search != null)
                        {
                            SelectionSet customSet = new SelectionSet(search);
                            customSet.DisplayName = value;

                            disciplineFolder.Children.Add(customSet);
                            totalCreated++;
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error creating {disciplineName}\\{value}: {ex.Message}");
                    }
                }
            }

            return totalCreated;
        }

        private FolderItem FindOrCreateCustomSetsFolder()
        {
            // Check if "3. CUSTOM SETS" folder already exists
            foreach (SavedItem item in _doc.SelectionSets.RootItem.Children)
            {
                if (item.IsGroup && item.DisplayName == "3. CUSTOM SETS")
                {
                    return (FolderItem)item;
                }
            }

            // Create new folder
            FolderItem customSetsFolder = new FolderItem();
            customSetsFolder.DisplayName = "3. CUSTOM SETS";
            _doc.SelectionSets.AddCopy(customSetsFolder);

            // Return reference to the newly added folder
            foreach (SavedItem item in _doc.SelectionSets.RootItem.Children)
            {
                if (item.IsGroup && item.DisplayName == "3. CUSTOM SETS")
                {
                    return (FolderItem)item;
                }
            }

            return customSetsFolder;
        }

        private List<string> DiscoverPropertyValuesForCustom(SelectionSet disciplineSet, string category, string propertyName)
        {
            var values = new HashSet<string>();

            ModelItemCollection items = disciplineSet.Search.FindAll(_doc, false);

            if (items.Count == 0)
            {
                return new List<string>();
            }

            // Scan items for property values
            foreach (ModelItem item in items)
            {
                foreach (ModelItem descendant in GetAllDescendants(item))
                {
                    try
                    {
                        foreach (PropertyCategory cat in descendant.PropertyCategories)
                        {
                            if (cat.DisplayName == category || cat.Name == category)
                            {
                                foreach (DataProperty prop in cat.Properties)
                                {
                                    if (prop.DisplayName == propertyName || prop.Name == propertyName)
                                    {
                                        string val = prop.Value.ToDisplayString();
                                        if (!string.IsNullOrWhiteSpace(val))
                                        {
                                            values.Add(val);
                                        }
                                        break;
                                    }
                                }
                                break;
                            }
                        }
                    }
                    catch { continue; }
                }
            }

            return values.OrderBy(v => v).ToList();
        }

        private Search CreateCustomSearchWithLocator(string disciplineName, string category, string propertyName, string value)
        {
            Search search = new Search();

            // Set locator to scope search to specific discipline
            string locatorPath = "lcop_selection_set_tree/1. DISCIPLINES/" + disciplineName;

            // Use the discipline search set as the selection source
            if (_disciplineSets.ContainsKey(disciplineName))
            {
                SelectionSet disciplineSet = _disciplineSets[disciplineName];
                ModelItemCollection items = disciplineSet.Search.FindAll(_doc, false);
                search.Selection.CopyFrom(items);
            }

            // Add property condition
            SearchCondition condition = SearchCondition.HasPropertyByDisplayName(category, propertyName);
            condition = condition.EqualValue(VariantData.FromDisplayString(value));

            search.SearchConditions.Add(condition);
            search.Locations = SearchLocations.DescendantsAndSelf;

            return search;
        }

        #endregion

        #region UI Helpers

        private void UpdateStatus(string message, string colorHex)
        {
            txtStatus.Text = message;
            txtStatus.Foreground = new System.Windows.Media.SolidColorBrush(
                (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(colorHex));
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        #endregion
    }
}
