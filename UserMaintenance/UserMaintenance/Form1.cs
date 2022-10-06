using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
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

            label1.Text = Resource1.LastName;
            label2.Text = Resource1.FirstName;
            button1.Text = Resource1.Add;

            //kijeloljuk a listbox forrasat (a user nevu binding list)
            listBox1.DataSource = users;
            //kijeloljuk hogy a forrasnak melyik parameteret mutassa a lista
            listBox1.DisplayMember = "FullName";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var u = new User()
            {
                LastName = textBox1.Text,
                FirstName = textBox2.Text
            };
            users.Add(u);

            textBox1.Clear();
            textBox2.Clear();
        }
    }
}
