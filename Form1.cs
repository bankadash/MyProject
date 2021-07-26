using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Calculator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnSumm_Click(object sender, EventArgs e)
        {
            IPizzaReciepe _icl=new EuropePizza();
            label1.Text = _icl.SetPizzaWeight().ToString();
        }

        private void btnDiff_Click(object sender, EventArgs e)
        {
            IPizzaReciepe _icl = new AmericanPizza();
            label1.Text = _icl.SetPizzaWeight().ToString();
        }
    }

     
}
