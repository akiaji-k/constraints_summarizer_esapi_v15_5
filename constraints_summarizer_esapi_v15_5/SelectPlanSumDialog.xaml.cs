using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace constraints_summarizer_esapi_v15_5
{
    /// <summary>
    /// Interaction logic for SelectPlanSumDialog.xaml
    /// </summary>
    public partial class SelectPlanSumDialog : Window
    {
        public string SelectedItem { get; private set; }
        public SelectPlanSumDialog(string message, List<string> options)
        {
            InitializeComponent();

            MessageLabel.Content = message;

            if (options.Count < 1)
            {
                throw new Exception("制約を確認するPlanの候補がありません(プランを開いてください)。");
            }
            else
            {
                foreach(var option in options)
                {
                    var item = new ListBoxItem();
                    item.Content = option;
                    item.MouseUp += ItemMouseUp;
                    listbox.Items.Add(item);
                }
            }
        }

        private void ItemMouseUp(object sender, MouseButtonEventArgs e)
        {
            var selected_item = (ListBoxItem)sender;
//            SelectedItem = (string)listbox.SelectedItem;
            SelectedItem = selected_item.Content.ToString();
            DialogResult = true;
            Close();

            return;
        }
    }
}
