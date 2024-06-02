using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutomaticMailPrinter
{
    public partial class FormOrders : Form
    {
        Database database = new Database();

        public FormOrders()
        {
            InitializeComponent();
        }

        private void FormOrders_Load(object sender, EventArgs e)
        {
            DataTable orders = database.GetOrders(10, 1);

            foreach (DataRow dr in orders.Rows)
            {
                ListViewItem item = new ListViewItem(dr["id"].ToString());
                item.Tag = dr["id"].ToString();
                item.SubItems.Add(dr["created_at"].ToString());
                item.SubItems.Add(dr["subject"].ToString());
                listViewOrders.Items.Add(item);
            }
        }

        private void buttonPrint_Click(object sender, EventArgs e)
        {
            try { 
                if (listViewOrders.SelectedItems.Count > 0) {
                    buttonPrint.Enabled = false;
                    buttonPrint.Text = "Printing...";
                    int id = int.Parse((string)listViewOrders.SelectedItems[0].Tag);
                    string html = database.GetOrder(id).html;
                    Program.PrintHtmlPage(html);
                    database.OrderPrinted(id);
                    buttonPrint.Text = "Print";
                    buttonPrint.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to print order. Please check the log for more details.", "Print failed!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Logger.LogError("Failed to re-print order", ex);
            }
        }
    }
}
