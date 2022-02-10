using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Resources;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DataCollector
{
    public partial class input : Form
    {
        public input()
        {
            InitializeComponent();
        }

        private void input_Load(object sender, EventArgs e)
        {
            var ResxFile = "SheetResource.resx";
            string sheetid = "";
            string cases = "";
            using (var reader = new ResXResourceReader(ResxFile))
            {
                foreach (DictionaryEntry d in reader)
                {
                    if (d.Key.ToString().Equals("SHEET_ID"))
                    {
                        sheetid = d.Value.ToString();
                    }
                    if (d.Key.ToString().Equals("CASES"))
                    {
                        cases = d.Value.ToString();
                    }
                }
            }
            this.sheetid.Text = sheetid;
            this.cases.Text = cases;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var ResxFile = "SheetResource.resx";
            using (var writer = new ResXResourceWriter(ResxFile))
            {
                writer.AddResource("SHEET_ID", sheetid.Text);
                writer.AddResource("CASES", cases.Text);
                writer.Generate();
                writer.Close();
            }
            this.Close();
        }
    }
}
