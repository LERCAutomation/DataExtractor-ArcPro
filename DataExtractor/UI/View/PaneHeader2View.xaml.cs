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

using System.Collections.Generic;
using System.Linq;
using System.Windows.Controls;

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

        private void ListViewPartners_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            List<Partner> added = e.AddedItems.OfType<Partner>().ToList();
            List<Partner> removed = e.RemovedItems.OfType<Partner>().ToList();

            var listView = sender as ListView;
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
        private void ListViewSQLLayers_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            List<SQLLayer> added = e.AddedItems.OfType<SQLLayer>().ToList();
            List<SQLLayer> removed = e.RemovedItems.OfType<SQLLayer>().ToList();

            var listView = sender as ListView;
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
        private void ListViewMapLayers_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            List<MapLayer> added = e.AddedItems.OfType<MapLayer>().ToList();
            List<MapLayer> removed = e.RemovedItems.OfType<MapLayer>().ToList();

            var listView = sender as ListView;
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
    }
}