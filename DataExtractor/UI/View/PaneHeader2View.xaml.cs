// The DataTools are a suite of ArcGIS Pro addins used to extract
// and manage biodiversity information from ArcGIS Pro and SQL Server
// based on pre-defined or user specified criteria.
//
// Copyright © 2024 Andy Foy Consulting.
//
// This file is part of DataTools suite of programs..
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

using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using MessageBox = ArcGIS.Desktop.Framework.Dialogs.MessageBox;

namespace DataExtractor.UI
{
    /// <summary>
    /// Interaction logic for PaneHeader2View.xaml
    /// </summary>
    public partial class PaneHeader2View : System.Windows.Controls.UserControl
    {
        public PaneHeader2View()
        {
            InitializeComponent();
        }

        private void ListViewPartners_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            List<Partner> added = e.AddedItems.OfType<Partner>().ToList();
            List<Partner> removed = e.RemovedItems.OfType<Partner>().ToList();

            var listView = sender as System.Windows.Controls.ListView;
            var itemsSelected = listView.Items.OfType<Partner>().ToList().Where(s => s.IsSelected == true).ToList();
            var itemsUnselected = listView.Items.OfType<Partner>().Where(p => p.IsSelected == false).ToList();
            var selectedItems = listView.SelectedItems.OfType<Partner>().ToList();

            // Ensure any removed items are actually unselected.
            if (removed.Count > 1)
            {
                e.RemovedItems.OfType<Partner>().ToList().ForEach(p => p.IsSelected = false);

                if (selectedItems.Count == 1)
                    listView.Items.OfType<Partner>().ToList().Where(s => selectedItems.All(s2 => s2.PartnerName != s.PartnerName)).ToList().ForEach(p => p.IsSelected = false);
            }
        }

        private void ListViewPartners_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var partner = ListViewPartners.SelectedItem as Partner;
            if (partner != null)
            {
                string notes = (string.IsNullOrEmpty(partner.Notes.Trim()) ? string.Empty : "\r\n\r\nNotes : " + partner.Notes);

                // Display the selected partner's details.
                string strText = string.Format("{0} ({1})\r\nGIS Format : {2}\r\nExport Format : {3}\r\n\r\nSQL Table : {4}\r\nSQL Files : {5}\r\n\r\nMap Files : {6}{7}",
                    partner.PartnerName, partner.ShortName, partner.GISFormat, partner.ExportFormat, partner.SQLTable, partner.SQLFiles.Replace(" ","").Replace(",", ", "), partner.MapFiles.Replace(" ", "").Replace(",", ", "), notes);
                MessageBox.Show(strText, "Partner Details", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void ListViewSQLLayers_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            List<SQLLayer> added = e.AddedItems.OfType<SQLLayer>().ToList();
            List<SQLLayer> removed = e.RemovedItems.OfType<SQLLayer>().ToList();

            var listView = sender as System.Windows.Controls.ListView;
            var itemsSelected = listView.Items.OfType<SQLLayer>().ToList().Where(s => s.IsSelected == true).ToList();
            var itemsUnselected = listView.Items.OfType<SQLLayer>().Where(p => p.IsSelected == false).ToList();
            var selectedItems = listView.SelectedItems.OfType<SQLLayer>().ToList();

            // Ensure any removed items are actually unselected.
            if (removed.Count > 1)
            {
                e.RemovedItems.OfType<SQLLayer>().ToList().ForEach(p => p.IsSelected = false);

                if (selectedItems.Count == 1)
                    listView.Items.OfType<SQLLayer>().ToList().Where(s => selectedItems.All(s2 => s2.NodeName != s.NodeName)).ToList().ForEach(p => p.IsSelected = false);
            }
        }

        private void ListViewSQLLayers_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var sqlLayer = ListViewSQLLayers.SelectedItem as SQLLayer;
            if (sqlLayer != null)
            {
                string outputType = (string.IsNullOrEmpty(sqlLayer.OutputType) ? string.Empty : "\r\nOutput Type : " + sqlLayer.OutputType);
                string macroName = (string.IsNullOrEmpty(sqlLayer.MacroName) ? string.Empty : "\r\n\r\nMacro Name : " + sqlLayer.MacroName);
                string macroParms = (string.IsNullOrEmpty(sqlLayer.MacroParms) ? string.Empty : "\r\n\r\nMacro Parms : " + sqlLayer.MacroParms);
                string whereClause = (string.IsNullOrEmpty(sqlLayer.WhereClause) ? string.Empty : "\r\n\r\nWhere Clause : " + sqlLayer.WhereClause);
                string orderBy = (string.IsNullOrEmpty(sqlLayer.OrderColumns) ? string.Empty : "\r\n\r\nOrder By : " + sqlLayer.OrderColumns);

                // Display the selected sql layer's details.
                string strText = string.Format("{0}\r\nOutput Name : {1}{2}{3}{4}{5}{6}",
                    sqlLayer.NodeName, sqlLayer.OutputName, outputType, whereClause, orderBy, macroName, macroParms);
                MessageBox.Show(strText, "SQL Layer Details", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void ListViewMapLayers_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            List<MapLayer> added = e.AddedItems.OfType<MapLayer>().ToList();
            List<MapLayer> removed = e.RemovedItems.OfType<MapLayer>().ToList();

            var listView = sender as System.Windows.Controls.ListView;
            var itemsSelected = listView.Items.OfType<MapLayer>().ToList().Where(s => s.IsSelected == true).ToList();
            var itemsUnselected = listView.Items.OfType<MapLayer>().Where(p => p.IsSelected == false).ToList();
            var selectedItems = listView.SelectedItems.OfType<MapLayer>().ToList();

            // Ensure any removed items are actually unselected.
            if (removed.Count > 1)
            {
                e.RemovedItems.OfType<MapLayer>().ToList().ForEach(p => p.IsSelected = false);

                if (selectedItems.Count == 1)
                    listView.Items.OfType<MapLayer>().ToList().Where(s => selectedItems.All(s2 => s2.NodeName != s.NodeName)).ToList().ForEach(p => p.IsSelected = false);
            }
        }

        private void ListViewMapLayers_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var mapLayer = ListViewMapLayers.SelectedItem as MapLayer;
            if (mapLayer != null)
            {
                string outputType = (string.IsNullOrEmpty(mapLayer.OutputType) ? string.Empty : "\r\nOutput Type : " + mapLayer.OutputType);
                string macroName = (string.IsNullOrEmpty(mapLayer.MacroName) ? string.Empty : "\r\n\r\nMacro Name : " + mapLayer.MacroName);
                string macroParms = (string.IsNullOrEmpty(mapLayer.MacroParms) ? string.Empty : "\r\n\r\nMacro Parms : " + mapLayer.MacroParms);
                string whereClause = (string.IsNullOrEmpty(mapLayer.WhereClause) ? string.Empty : "\r\n\r\nWhere Clause : " + mapLayer.WhereClause);
                string orderBy = (string.IsNullOrEmpty(mapLayer.OrderColumns) ? string.Empty : "\r\n\r\nOrder By : " + mapLayer.OrderColumns);

                // Display the selected sql layer's details.
                string strText = string.Format("{0}\r\nLayer Name : {1}\r\nOutput Name : {2}{3}{4}{5}{6}{7}",
                    mapLayer.NodeName, mapLayer.LayerName, mapLayer.OutputName, outputType, whereClause, orderBy, macroName, macroParms);
                MessageBox.Show(strText, "Map Layer Details", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void ListViewPartners_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            var listView = sender as System.Windows.Controls.ListView;

            ScrollViewer sv = FindVisualChild<ScrollViewer>(listView);
            Visibility vsVisibility = sv.ComputedVerticalScrollBarVisibility;
            double vsWidth = ((vsVisibility == Visibility.Visible) ? SystemParameters.VerticalScrollBarWidth : 0);

            var gridView = listView.View as GridView;
            gridView.Columns[0].Width = listView.ActualWidth - vsWidth - 10;
        }

        private void ListViewSQLLayers_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            var listView = sender as System.Windows.Controls.ListView;

            ScrollViewer sv = FindVisualChild<ScrollViewer>(listView);
            Visibility vsVisibility = sv.ComputedVerticalScrollBarVisibility;
            double vsWidth = ((vsVisibility == Visibility.Visible) ? SystemParameters.VerticalScrollBarWidth : 0);

            var gridView = listView.View as GridView;
            gridView.Columns[0].Width = listView.ActualWidth - vsWidth - 10;
        }

        private void ListViewMapLayers_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            var listView = sender as System.Windows.Controls.ListView;

            ScrollViewer sv = FindVisualChild<ScrollViewer>(listView);
            Visibility vsVisibility = sv.ComputedVerticalScrollBarVisibility;
            double vsWidth = ((vsVisibility == Visibility.Visible) ? SystemParameters.VerticalScrollBarWidth : 0);

            var gridView = listView.View as GridView;
            gridView.Columns[0].Width = listView.ActualWidth - vsWidth - 10;
        }

        private childItem FindVisualChild<childItem>(DependencyObject obj)
               where childItem : DependencyObject

        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(obj); i++)
            {
                DependencyObject child = VisualTreeHelper.GetChild(obj, i);
                if (child != null && child is childItem)
                    return (childItem)child;
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