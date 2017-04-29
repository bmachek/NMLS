using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace Editor
{
    public partial class EditorProgram : Form
    {
        public EditorProgram()
        {
            InitializeComponent();

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            label30.Text = Program.directoryEntry.Name;

            textBoxes[1] = textBox1;
            textBoxes[2] = textBox2;
            textBoxes[3] = textBox3;
            textBoxes[4] = textBox4;
            textBoxes[5] = textBox5;
            textBoxes[6] = textBox6;
            textBoxes[7] = textBox7;
            textBoxes[8] = textBox8;
            textBoxes[9] = textBox9;
            textBoxes[10] = textBox10;
            textBoxes[11] = textBox11;
            textBoxes[12] = textBox12;
            textBoxes[13] = textBox13;
            textBoxes[14] = textBox14;
            textBoxes[15] = textBox15;
            textBoxes[16] = textBox16;
            textBoxes[17] = textBox17;
            textBoxes[18] = textBox18;
            textBoxes[19] = textBox19;
            textBoxes[20] = textBox20;
            textBoxes[21] = textBox21;
            textBoxes[22] = textBox22;
            textBoxes[23] = textBox23;
            textBoxes[24] = textBox24;
            textBoxes[25] = textBox25;
            textBoxes[26] = textBox26;

            checkBoxes[1] = checkBox1;
            checkBoxes[2] = checkBox2;
            checkBoxes[3] = checkBox3;
            checkBoxes[4] = checkBox4;
            checkBoxes[5] = checkBox5;
            checkBoxes[6] = checkBox6;
            checkBoxes[7] = checkBox7;
            checkBoxes[8] = checkBox8;
            checkBoxes[9] = checkBox9;
            checkBoxes[10] = checkBox10;
            checkBoxes[11] = checkBox11;
            checkBoxes[12] = checkBox12;
            checkBoxes[13] = checkBox13;
            checkBoxes[14] = checkBox14;
            checkBoxes[15] = checkBox15;
            checkBoxes[16] = checkBox16;
            checkBoxes[17] = checkBox17;
            checkBoxes[18] = checkBox18;
            checkBoxes[19] = checkBox19;
            checkBoxes[20] = checkBox20;
            checkBoxes[21] = checkBox21;
            checkBoxes[22] = checkBox22;
            checkBoxes[23] = checkBox23;
            checkBoxes[24] = checkBox24;
            checkBoxes[25] = checkBox25;
            checkBoxes[26] = checkBox26;

            labels[1] = uncPath1;
            labels[2] = uncPath2;
            labels[3] = uncPath3;
            labels[4] = uncPath4;
            labels[5] = uncPath5;
            labels[6] = uncPath6;
            labels[7] = uncPath7;
            labels[8] = uncPath8;
            labels[9] = uncPath9;
            labels[10] = uncPath10;
            labels[11] = uncPath11;
            labels[12] = uncPath12;
            labels[13] = uncPath13;
            labels[14] = uncPath14;
            labels[15] = uncPath15;
            labels[16] = uncPath16;
            labels[17] = uncPath17;
            labels[18] = uncPath18;
            labels[19] = uncPath19;
            labels[20] = uncPath20;
            labels[21] = uncPath21;
            labels[22] = uncPath22;
            labels[23] = uncPath23;
            labels[24] = uncPath24;
            labels[25] = uncPath25;
            labels[26] = uncPath26;


            LoadFromLDAP();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void label30_Click(object sender, EventArgs e)
        {

        }

        private void checkBox27_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void textBox27_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            SaveToLDAP();
            Dispose();
        }

        private void SaveToLDAP()
        {

            for (int i = 1; i <= 26; i++)
            {
                char driveLetter = Convert.ToChar(i + 64);

                if (textBoxes[i].Text != null)
                {
                    if (textBoxes[i].Text.Equals(""))
                    {
                        Program.directoryEntry.Properties["nmlsNetworkDrive" + driveLetter].Clear();
                    }
                    else
                    {
                        Program.directoryEntry.Properties["nmlsNetworkDrive" + driveLetter].Value = textBoxes[i].Text;
                    }
                }

                Program.directoryEntry.Properties["nmlsNetworkDriveDel" + driveLetter].Value = checkBoxes[i].Checked;
            }

            Program.directoryEntry.CommitChanges();


        }


        private void LoadFromLDAP()
        {
            for (int i = 1; i <= 26; i++)
            {
                char driveLetter = Convert.ToChar(i + 64);

                if (Program.directoryEntry.Properties["nmlsNetworkDrive" + driveLetter].Value != null)
                {
                    textBoxes[i].Text = Program.directoryEntry.Properties["nmlsNetworkDrive" + driveLetter].Value.ToString();
                    labels[i].Dispose();
                }

                if (Program.directoryEntry.Properties["nmlsNetworkDriveDel" + driveLetter].Value != null)
                {
                    object delete = Program.directoryEntry.Properties["nmlsNetworkDriveDel" + driveLetter].Value;
                    if (((Boolean)delete))
                    {
                        checkBoxes[i].Checked = true;
                    }
                    else
                    {
                        checkBoxes[i].Checked = false;
                    }

                }
            }

            for (int i = 1; i < 27; i++)
            {
                char driveLetter = Convert.ToChar(i + 64);

                // System.Windows.Forms.MessageBox.Show(driveLetter + "");

                if (Program.nmlsPathes.ContainsKey(driveLetter))
                {
                    labels[i].Text = Program.nmlsPathes[driveLetter];
                    labels[i].BackColor = Color.Transparent;
                }
                else
                {
                    labels[i].Dispose();

                }

                if (Program.nmlsDelDrives.ContainsKey(driveLetter))
                {
                    if (!checkBoxes[i].Checked)
                    {
                        checkBoxes[i].Checked = Program.nmlsDelDrives[driveLetter];
                    }
                }
            }
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }
    }
}
