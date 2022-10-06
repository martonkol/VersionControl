using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using UserMaintenance.Entities;

namespace UserMaintenance
{
    public partial class Form1 : Form
    {
        //létrehozok egy User(.cs) típusú BindingListet - kell a using UserMaintenance.Entities
        BindingList<User> users = new BindingList<User>();

        //a Form1 konstruktora
        public Form1()
        {
            InitializeComponent();

            label1.Text = Resource1.FullName;
            button1.Text = Resource1.Add;
            button2.Text = Resource1.Write;

            //kijeloljuk a listbox forrasat (a user nevu binding list)
            listBox1.DataSource = users;
            //kijeloljuk hogy a forrasnak melyik parameteret mutassa a lista
            listBox1.DisplayMember = "FullName";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var u = new User()
            {
                FullName = textBox1.Text
            };
            users.Add(u);

            textBox1.Clear();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SaveFileDialog sFD = new SaveFileDialog();
            sFD.InitialDirectory = Application.StartupPath;
            sFD.Filter = "Vesszővel tagolt értékek (*.csv)|*.csv";
            sFD.DefaultExt = "csv";
            sFD.AddExtension = true;

            if (sFD.ShowDialog() == DialogResult.OK)
            {
                StreamWriter sw = new StreamWriter(sFD.FileName, false, Encoding.Default);

                sw.WriteLine("ID;FullName");

                foreach (User item in users)
                {
                    sw.WriteLine($"{item.ID};{item.FullName}");
                }
                sw.Close();

            }
        }
    }
}
