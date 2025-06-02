// The DataTools are a suite of ArcGIS Pro addins used to extract
// and manage biodiversity information from ArcGIS Pro and SQL Server
// based on pre-defined or user specified criteria.
//
// Copyright © 2024-25 Andy Foy Consulting.
//
// This file is part of DataTools suite of programs.
//
// DataTools are free software: you can redistribute it and/or modify
// them under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// DataTools are distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with with program.  If not, see <http://www.gnu.org/licenses/>.

using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using MessageBox = ArcGIS.Desktop.Framework.Dialogs.MessageBox;

namespace DataExtractor.UI
{
    /// <summary>
    /// Interaction logic for PaneHeader2View.xaml
    /// </summary>
    public partial class PaneHeader2View : UserControl
    {
        public PaneHeader2View()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Ensure any removed partners are actually unselected.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ListViewPartners_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Get the list of removed items.
            List<Partner> removed = e.RemovedItems.OfType<Partner>().ToList();

            // Ensure any removed items are actually unselected.
            if (removed.Count > 1)
            {
                // Unselect the removed items.
                e.RemovedItems.OfType<Partner>().ToList().ForEach(p => p.IsSelected = false);

                // Get the list of currently selected items.
                var listView = sender as System.Windows.Controls.ListView;
                var selectedItems = listView.SelectedItems.OfType<Partner>().ToList();

                if (selectedItems.Count == 1)
                    listView.Items.OfType<Partner>().ToList().Where(s => selectedItems.All(s2 => s2.PartnerName != s.PartnerName)).ToList().ForEach(p => p.IsSelected = false);
            }
        }

        /// <summary>
        /// Display the details when a partner is double-clicked.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ListViewPartners_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            // Get the original element that was double-clicked on
            // and search from child to parent until you find either
            // a ListViewItem or the top of the tree.
            DependencyObject originalSource = (DependencyObject)e.OriginalSource;
            while ((originalSource != null) && originalSource is not System.Windows.Controls.ListViewItem)
            {
                originalSource = VisualTreeHelper.GetParent(originalSource);
            }

            // If it didn’t find a ListViewItem anywhere in the hierarchy
            // then it’s because the user didn’t click on one. Therefore
            // if the variable isn’t null, run the code.
            if (originalSource != null)
            {
                if (ListViewPartners.SelectedItem is Partner partner)
                {
                    string notes = (string.IsNullOrEmpty(partner.Notes.Trim()) ? string.Empty : "\r\n\r\nNotes : " + partner.Notes);

                    // Display the selected partner's details.
                    string strText = string.Format("{0} ({1})\r\nGIS Format : {2}\r\nExport Format : {3}\r\n\r\nSQL Table : {4}\r\nSQL Files : {5}\r\n\r\nMap Files : {6}{7}",
                        partner.PartnerName, partner.ShortName, partner.GISFormat, partner.ExportFormat, partner.SQLTable, partner.SQLFiles.Replace(" ", "").Replace(",", ", "), partner.MapFiles.Replace(" ", "").Replace(",", ", "), notes);
                    MessageBox.Show(strText, "Partner Details", MessageBoxButton.OK);
                }
            }
        }

        /// <summary>
        /// Override <Ctrl>A behaviour to make sure all list items are
        /// selected, not just the ones that are visible.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ListViewPartners_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if ((e.Key == Key.A) && (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl)))
            {
                foreach (Partner partner in ListViewPartners.Items)
                {
                    partner.IsSelected = true;
                }
                e.Handled = true;
            }
        }

        /// <summary>
        /// Ensure any removed SQL layers are actually unselected.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ListViewSQLLayers_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Get the list of removed items.
            List<SQLLayer> removed = e.RemovedItems.OfType<SQLLayer>().ToList();

            // Ensure any removed items are actually unselected.
            if (removed.Count > 1)
            {
                // Unselect the removed items.
                e.RemovedItems.OfType<SQLLayer>().ToList().ForEach(p => p.IsSelected = false);

                // Get the list of currently selected items.
                var listView = sender as System.Windows.Controls.ListView;
                var selectedItems = listView.SelectedItems.OfType<SQLLayer>().ToList();

                if (selectedItems.Count == 1)
                    listView.Items.OfType<SQLLayer>().ToList().Where(s => selectedItems.All(s2 => s2.NodeName != s.NodeName)).ToList().ForEach(p => p.IsSelected = false);
            }
        }

        /// <summary>
        /// Display the details when a SQL layer is double-clicked.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ListViewSQLLayers_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            // Get the original element that was double-clicked on
            // and search from child to parent until you find either
            // a ListViewItem or the top of the tree.
            DependencyObject originalSource = (DependencyObject)e.OriginalSource;
            while ((originalSource != null) && originalSource is not System.Windows.Controls.ListViewItem)
            {
                originalSource = VisualTreeHelper.GetParent(originalSource);
            }

            // If it didn’t find a ListViewItem anywhere in the hierarchy
            // then it’s because the user didn’t click on one. Therefore
            // if the variable isn’t null, run the code.
            if (originalSource != null)
            {
                if (ListViewSQLLayers.SelectedItem is SQLLayer sqlLayer)
                {
                    string outputType = (string.IsNullOrEmpty(sqlLayer.OutputType) ? string.Empty : "\r\nOutput Type : " + sqlLayer.OutputType);
                    string macroName = (string.IsNullOrEmpty(sqlLayer.MacroName) ? string.Empty : "\r\n\r\nMacro Name : " + sqlLayer.MacroName);
                    string macroParms = (string.IsNullOrEmpty(sqlLayer.MacroParms) ? string.Empty : "\r\n\r\nMacro Parms : " + sqlLayer.MacroParms);
                    string whereClause = (string.IsNullOrEmpty(sqlLayer.WhereClause) ? string.Empty : "\r\n\r\nWhere Clause : " + sqlLayer.WhereClause);
                    string orderBy = (string.IsNullOrEmpty(sqlLayer.OrderColumns) ? string.Empty : "\r\n\r\nOrder By : " + sqlLayer.OrderColumns);

                    // Display the selected SQL layer's details.
                    string strText = string.Format("{0}\r\nOutput Name : {1}{2}{3}{4}{5}{6}",
                        sqlLayer.NodeName, sqlLayer.OutputName, outputType, whereClause, orderBy, macroName, macroParms);
                    MessageBox.Show(strText, "SQL Layer Details", MessageBoxButton.OK);
                }
            }
        }

        /// <summary>
        /// Override <Ctrl>A behaviour to make sure all list items are
        /// selected, not just the ones that are visible.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ListViewSQLLayers_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if ((e.Key == Key.A) && (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl)))
            {
                foreach (SQLLayer sqlLayer in ListViewSQLLayers.Items)
                {
                    sqlLayer.IsSelected = true;
                }
                e.Handled = true;
            }
        }

        /// <summary>
        /// Ensure any removed map layers are actually unselected.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ListViewMapLayers_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Get the list of removed items.
            List<MapLayer> removed = e.RemovedItems.OfType<MapLayer>().ToList();

            // Ensure any removed items are actually unselected.
            if (removed.Count > 1)
            {
                // Unselect the removed items.
                e.RemovedItems.OfType<MapLayer>().ToList().ForEach(p => p.IsSelected = false);

                // Get the list of currently selected items.
                var listView = sender as System.Windows.Controls.ListView;
                var selectedItems = listView.SelectedItems.OfType<MapLayer>().ToList();

                if (selectedItems.Count == 1)
                    listView.Items.OfType<MapLayer>().ToList().Where(s => selectedItems.All(s2 => s2.NodeName != s.NodeName)).ToList().ForEach(p => p.IsSelected = false);
            }
        }

        /// <summary>
        /// Display the details when a map layer is double-clicked.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ListViewMapLayers_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            // Get the original element that was double-clicked on
            // and search from child to parent until you find either
            // a ListViewItem or the top of the tree.
            DependencyObject originalSource = (DependencyObject)e.OriginalSource;
            while ((originalSource != null) && originalSource is not System.Windows.Controls.ListViewItem)
            {
                originalSource = VisualTreeHelper.GetParent(originalSource);
            }

            // If it didn’t find a ListViewItem anywhere in the hierarchy
            // then it’s because the user didn’t click on one. Therefore
            // if the variable isn’t null, run the code.
            if (originalSource != null)
            {
                if (ListViewMapLayers.SelectedItem is MapLayer mapLayer)
                {
                    string outputType = (string.IsNullOrEmpty(mapLayer.OutputType) ? string.Empty : "\r\nOutput Type : " + mapLayer.OutputType);
                    string macroName = (string.IsNullOrEmpty(mapLayer.MacroName) ? string.Empty : "\r\n\r\nMacro Name : " + mapLayer.MacroName);
                    string macroParms = (string.IsNullOrEmpty(mapLayer.MacroParms) ? string.Empty : "\r\n\r\nMacro Parms : " + mapLayer.MacroParms);
                    string whereClause = (string.IsNullOrEmpty(mapLayer.WhereClause) ? string.Empty : "\r\n\r\nWhere Clause : " + mapLayer.WhereClause);
                    string orderBy = (string.IsNullOrEmpty(mapLayer.OrderColumns) ? string.Empty : "\r\n\r\nOrder By : " + mapLayer.OrderColumns);

                    // Display the selected map layer's details.
                    string strText = string.Format("{0}\r\nLayer Name : {1}\r\nOutput Name : {2}{3}{4}{5}{6}{7}",
                        mapLayer.NodeName, mapLayer.LayerName, mapLayer.OutputName, outputType, whereClause, orderBy, macroName, macroParms);
                    MessageBox.Show(strText, "Map Layer Details", MessageBoxButton.OK);
                }
            }
        }

        /// <summary>
        /// Override <Ctrl>A behaviour to make sure all list items are
        /// selected, not just the ones that are visible.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ListViewMapLayers_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if ((e.Key == Key.A) && (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl)))
            {
                foreach (MapLayer mapLayer in ListViewMapLayers.Items)
                {
                    mapLayer.IsSelected = true;
                }
                e.Handled = true;
            }
        }

        /// <summary>
        /// Reset the width of the partner column to match the width of the list view.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ListViewPartners_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            // Get the partner list view.
            var listView = sender as System.Windows.Controls.ListView;

            // Get the scroll viewer for the list and check if it is visible.
            ScrollViewer sv = FindVisualChild<ScrollViewer>(listView);
            Visibility vsVisibility = sv.ComputedVerticalScrollBarVisibility;

            // Calculate the width of the scroll viewer.
            double vsWidth = ((vsVisibility == Visibility.Visible) ? SystemParameters.VerticalScrollBarWidth : 0);

            // Set the width of the first column in the list view.
            var gridView = listView.View as GridView;
            gridView.Columns[0].Width = listView.ActualWidth - vsWidth - 10;
        }

        /// <summary>
        /// Reset the width of the SQL layer column to match the width of the list view.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ListViewSQLLayers_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            // Get the SQL layer list view.
            var listView = sender as System.Windows.Controls.ListView;

            // Get the scroll viewer for the list and check if it is visible.
            ScrollViewer sv = FindVisualChild<ScrollViewer>(listView);
            Visibility vsVisibility = sv.ComputedVerticalScrollBarVisibility;

            // Calculate the width of the scroll viewer.
            double vsWidth = ((vsVisibility == Visibility.Visible) ? SystemParameters.VerticalScrollBarWidth : 0);

            // Set the width of the first column in the list view.
            var gridView = listView.View as GridView;
            gridView.Columns[0].Width = listView.ActualWidth - vsWidth - 10;
        }

        /// <summary>
        /// Reset the width of the map layer column to match the width of the list view.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ListViewMapLayers_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            // Get the SQL layer list view.
            var listView = sender as System.Windows.Controls.ListView;

            // Get the scroll viewer for the list and check if it is visible.
            ScrollViewer sv = FindVisualChild<ScrollViewer>(listView);
            Visibility vsVisibility = sv.ComputedVerticalScrollBarVisibility;

            // Calculate the width of the scroll viewer.
            double vsWidth = ((vsVisibility == Visibility.Visible) ? SystemParameters.VerticalScrollBarWidth : 0);

            // Set the width of the first column in the list view.
            var gridView = listView.View as GridView;
            gridView.Columns[0].Width = listView.ActualWidth - vsWidth - 10;
        }

        /// <summary>
        /// Return the first visual child object of the required type
        /// for the specified object.
        /// </summary>
        /// <typeparam name="childItem"></typeparam>
        /// <param name="obj"></param>
        /// <returns></returns>
        private static childItem FindVisualChild<childItem>(DependencyObject obj)
               where childItem : DependencyObject
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(obj); i++)
            {
                DependencyObject child = VisualTreeHelper.GetChild(obj, i);
                if (child != null && child is childItem item)
                    return item;
                else
                {
                    childItem childOfChild = FindVisualChild<childItem>(child);
                    if (childOfChild != null)
                        return childOfChild;
                }
            }

            return null;
        }
    }
}