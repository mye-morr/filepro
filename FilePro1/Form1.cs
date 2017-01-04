using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;

namespace FilePro1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        OpenFileDialog dlg_OpenFilel = new OpenFileDialog();
        OpenFileDialog dlg_OpenFile2 = new OpenFileDialog();
        OpenFileDialog dlg_OpenFile3 = new OpenFileDialog();
        OpenFileDialog dlg_OpenFile4 = new OpenFileDialog();
        FolderBrowserDialog dlg_BrowseFolder1 = new FolderBrowserDialog();

        private void checkStage()
        {
            int iStage = 0;

            if (btnBrowseInput.FlatStyle == FlatStyle.Flat)
            {
                iStage = 1;
            }
            else if (btnBig.FlatStyle == FlatStyle.Flat)
            {
                iStage = 2;
            }


            switch (iStage)
            {
                case 0:
                    break;
                case 1:
                    btnBrowseInput.Select();
                    break;
                case 2:
                    btnBig.Select();
                    break;
            }
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

            int i = 0;

            /*
            The ListView.SelectedIndexChanged event has a quirk that bombs your code.
            When you start your program, no item is selected. Click an item and SelectedIndexChanged fires, no problem.
            Now click another item and the event fires twice. First to let you know, unhelpfully, that the first item is unselected.
            Then again to tell you that the new item is selected. That first event is going to make you index an empty array, kaboom.
            RV1987's snippet prevents this.
            */

            if (listView1.SelectedIndices.Count > 0)
            {
                i = listView1.SelectedIndices[0];

                //TrackInformation t = (TrackInformation)SongList[listView1.SelectedIndices[0]];

                string sItem = listView1.Items[i].Text;

                txtboxDetails.Text = "";

                switch (sItem) { 
                    case "Demos( N)":
                        txtboxDetails.SelectionFont = new Font("Times New Roman", 10, FontStyle.Bold);
                        txtboxDetails.SelectedText = "Demos( N):\t[ ... {Acct} ... ]";
                        txtboxDetails.SelectionFont = new Font("Times New Roman", 10, FontStyle.Regular);
                        txtboxDetails.SelectedText = "\n\nEach demo worksheet must have a column called 'Acct', to denote patient identifiers, or the process will fail.";
                        txtboxDetails.SelectedText = "\n\nThe combined output from processing will be added to the right of the demographics data.";
                        break;
                    case "~Payments( N)":
                        txtboxDetails.SelectionFont = new Font("Times New Roman", 10, FontStyle.Bold);
                        txtboxDetails.SelectedText = "~Payments( N):\t[ {Acct}   {Insurer}   {Pmt} ]";
                        txtboxDetails.SelectionFont = new Font("Times New Roman", 10, FontStyle.Regular);
                        txtboxDetails.SelectedText = "\n\nOptional because the demographics file may already list out payors and payments, ordered by relevance / contribution.";
                        txtboxDetails.SelectedText = "\n\nThis worksheet may be left blank, although column headers should remain so as not to break the data-import process. This would result in no payments section being present in the output csv file(s).";
                        txtboxDetails.SelectedText = "\n\nAll additional N instances of this worksheet are combined together into one table and treated as a single unit for processing.";
                        break;
                    case "Rev Codes( N)":
                        txtboxDetails.SelectionFont = new Font("Times New Roman", 10, FontStyle.Bold);
                        txtboxDetails.SelectedText = "Rev Codes( N):\t[ {Acct}   {Rev}   {Chg} ]";
                        txtboxDetails.SelectionFont = new Font("Times New Roman", 10, FontStyle.Regular);
                        txtboxDetails.SelectedText = "\n\nThe magic of this sheet is that all charges for a given Acct and Revenue Code are combined, including positive and negative charges which may cancel out.";
                        txtboxDetails.SelectedText = "\n\nOutput section includes all revenue codes present in the input data.";
                        txtboxDetails.SelectedText = "\n\nAll additional N instances of this worksheet are combined together into one table and treated as a single unit for processing.";
                        break;
                    case "Diagnoses( N)":
                        txtboxDetails.SelectionFont = new Font("Times New Roman", 10, FontStyle.Bold);
                        txtboxDetails.SelectedText = "Diagnoses( N):\t[ {Acct}   {Dx}   {Seq} ]";
                        txtboxDetails.SelectionFont = new Font("Times New Roman", 10, FontStyle.Regular);
                        txtboxDetails.SelectedText = "\n\nFirst 40 diagnosis codes are ordered by sequence number in the following way: [ A, P, 1, 2, 3, ...]";
                        txtboxDetails.SelectedText = "\n\nIf need be, this behavior can be modified in the above SQL script file, using any text-editing program.";
                        txtboxDetails.SelectedText = "\n\nAll additional N instances of this worksheet are combined together into one table and treated as a single unit for processing.";
                        break;
                    case "~Procedures( N)":
                        txtboxDetails.SelectionFont = new Font("Times New Roman", 10, FontStyle.Bold);
                        txtboxDetails.SelectedText = "~Procedures( N):\t[ {Acct}   {Px}   {Seq} ]";
                        txtboxDetails.SelectionFont = new Font("Times New Roman", 10, FontStyle.Regular);
                        txtboxDetails.SelectedText = "\n\nOptional because the data may not have procedure codes.";
                        txtboxDetails.SelectedText = "\n\nThis worksheet may be left blank, although column headers should remain so as not to break the data-import process. This would result in no procedures section being present in the output csv file(s).";
                        txtboxDetails.SelectedText = "\n\nAll additional N instances of this worksheet are combined together into one table and treated as a single unit for processing.";
                        break;
                    case "Hcpcs( N)":
                        txtboxDetails.SelectionFont = new Font("Times New Roman", 10, FontStyle.Bold);
                        txtboxDetails.SelectedText = "Hcpcs( N):\t[ {Acct}   {HCPCS}   {Mod}   {Rev}   {Qty}   {Chg} ]";
                        txtboxDetails.SelectionFont = new Font("Times New Roman", 10, FontStyle.Regular);
                        txtboxDetails.SelectedText = "\n\nFirst 40 HCPCS codes, in order of appearance in the input data.";
                        txtboxDetails.SelectedText = "\n\nHCPCS data may not always be associated with Revenue data, but, currently, all 5 columns (HCPCS, Mod, Rev, Qty, Chg) will be present for each given HCPCS code. This presentation may change soon.";
                        txtboxDetails.SelectedText = "\n\nAll additional N instances of this worksheet are combined together into one table and treated as a single unit for processing.";
                        break;
                    case "Known Issues":
                        txtboxDetails.SelectionFont = new Font("Times New Roman", 10, FontStyle.Bold);
                        txtboxDetails.SelectedText = "Known Issues:";
                        txtboxDetails.SelectionFont = new Font("Times New Roman", 10, FontStyle.Regular);
                        txtboxDetails.SelectedText = "\n\nAt this time, there is no helpful error message if users forget to label the patient identifier column in the demographics data (as 'Acct').";
                        break;
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dlg_OpenFilel.Filter = "Excel 2010 Files|*.xlsx";
            dlg_OpenFile2.Filter = "SSIS Package|*.dtsx";
            dlg_OpenFile3.Filter = "Microsoft SQL Script|*.sql";
            dlg_OpenFile4.Filter = "Microsoft VBS Script|*.vbs";
            btnBrowseInput.FlatAppearance.BorderColor = Color.Orange;
            btnBrowseInput.FlatAppearance.BorderSize = 2;
            btnBrowseInput.FlatStyle = FlatStyle.Flat;

            listView1.Items[0].Selected = true;
            listView1.Select();
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            if (dlg_OpenFile2.ShowDialog() == DialogResult.OK)
            {
                txtboxImport.Text = dlg_OpenFile2.FileName;
            }

            checkStage();
        }

        private void btnProcess_Click(object sender, EventArgs e)
        {
            if (dlg_OpenFile3.ShowDialog() == DialogResult.OK)
            {
                txtboxProcess.Text = dlg_OpenFile3.FileName;
            }

            checkStage();
        }

        private void btnFormat_Click(object sender, EventArgs e)
        {
            if (dlg_OpenFile4.ShowDialog() == DialogResult.OK)
            {
                txtboxFormat.Text = dlg_OpenFile4.FileName;
            }

            checkStage();
        }

        private void btnBrowseOutput_Click(object sender, EventArgs e)
        {

            dlg_BrowseFolder1.SelectedPath = @"c:\input\";
            if (dlg_BrowseFolder1.ShowDialog() == DialogResult.OK)
            {
                String sPath = dlg_BrowseFolder1.SelectedPath;
                if (!String.IsNullOrWhiteSpace(sPath))
                {
                    txtboxOutput.Text = sPath;
                }
            }
        }

        private void btnBrowseInput_Click(object sender, EventArgs e)
        {
            if (dlg_OpenFilel.ShowDialog() == DialogResult.OK)
            {
                txtboxInput.Text = dlg_OpenFilel.FileName;

                if (btnBrowseInput.FlatStyle == FlatStyle.Flat)
                {
                    btnBrowseInput.FlatStyle = FlatStyle.Standard;

                    btnBig.FlatStyle = FlatStyle.Flat;
                    btnBig.FlatAppearance.BorderColor = Color.Orange;
                    btnBig.FlatAppearance.BorderSize = 2;
                    btnBig.Select();
                }
            }

            checkStage();
        }

        private void btnBig_Click(object sender, EventArgs e)
        {
            if (txtboxInput.Text.Equals("Your formatted *.xlsx file"))
                MessageBox.Show("You must supply a valid *.xlsx input file.", "Not so fast!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            else if (txtboxOutput.Text.Equals("Target folder for csv files"))
                MessageBox.Show("You must supply a valid folder for output.", "Not so fast!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            else
            {
                String text1 = File.ReadAllText(txtboxImport.Text);
                text1 = text1.Replace("***SOURCE_SERVER***", txtboxServer.Text);
                text1 = text1.Replace("***SOURCE_TEMPLATE***", txtboxInput.Text);

                String pathPackageFile = txtboxImport.Text.Replace("_params", "");
                File.WriteAllText(pathPackageFile, text1);

                String text2 = File.ReadAllText(txtboxProcess.Text);
                text2 = text2.Replace("***OUTPUT_PATH***", txtboxOutput.Text);
                text2 = text2.Replace("***SOURCE_SERVER***", txtboxServer.Text);

                String pathScriptFile = txtboxProcess.Text.Replace("_params", "");
                File.WriteAllText(pathScriptFile, text2);

                String text3 = File.ReadAllText(txtboxFormat.Text);
                text3 = text3.Replace("***SOURCE_TEMPLATE***", txtboxInput.Text);
                text3 = text3.Replace("***OUTPUT_PATH***", txtboxOutput.Text);

                String pathFormatFile = txtboxFormat.Text.Replace("_params", "");
                File.WriteAllText(pathFormatFile, text3);

                lblProcessing.Visible = true;

                String sCmd1 = 
                      "\"c:\\program files\\microsoft sql server\\130\\dts\\binn\\dtexec\" /FILE \"\\\""
                    + pathPackageFile
                    + "\\\"\" /CHECKPOINTING OFF /REPORTING EWCDI";

                String sCmd2 =
                      "sqlcmd -E -S \""
                    + txtboxServer.Text
                    + "\" -d demo_db -i \""
                    + pathScriptFile
                    + "\"";

                String sCmd3 = pathFormatFile;

                FileStream fs;
                fs = new FileStream(txtboxOutput.Text + "\\runme.bat", FileMode.Create, FileAccess.Write);
                StreamWriter writer = new StreamWriter(fs);
                writer.WriteLine(sCmd1);
                writer.WriteLine(sCmd2);
                writer.WriteLine(sCmd3);
                writer.Close();

                Process cmd0 = new Process();
                cmd0.StartInfo.FileName = "cmd.exe";
                cmd0.StartInfo.Arguments = "/C " + txtboxOutput.Text + "\\LiveLogViewer.exe\"";
                cmd0.StartInfo.RedirectStandardInput = true;
                cmd0.StartInfo.RedirectStandardOutput = true;
                cmd0.StartInfo.CreateNoWindow = true;
                cmd0.StartInfo.UseShellExecute = false;
                cmd0.Start();

                Process cmd1 = new Process();
                cmd1.StartInfo.FileName = "cmd.exe";
                cmd1.StartInfo.Arguments = "/C " + txtboxOutput.Text + "\\runme.bat > \"" + txtboxOutput.Text + "\\log.txt\"";
                cmd1.StartInfo.RedirectStandardInput = true;
                cmd1.StartInfo.RedirectStandardOutput = true;
                cmd1.StartInfo.CreateNoWindow = true;
                cmd1.StartInfo.UseShellExecute = false;

                cmd1.Start();
                cmd1.WaitForExit();
                
                try
                {
                    //File.Delete(pathScriptFile);
                    //File.Delete(pathPackageFile);
                    //File.Delete(pathFormatFile);
                    File.Delete(txtboxOutput.Text + "\\runme.bat");
                    File.Delete(txtboxOutput.Text + "\\output_1.csv");
                    File.Delete(txtboxOutput.Text + "\\output_2.csv");
                }
                catch (IOException) { }

                lblProcessing.Visible = false;

                // no longer necessary bc vbscript opens up result
                //MessageBox.Show("Thank you; come again.", "Process complete");
            }

            checkStage();
        }

        private void btnTest_Click(object sender, EventArgs e)
        {
            String text1 = File.ReadAllText(txtboxImport.Text);
            text1 = text1.Replace("***SOURCE_SERVER***", txtboxServer.Text);
            text1 = text1.Replace("***SOURCE_TEMPLATE***", txtboxInput.Text);

            String pathPackageFile = txtboxImport.Text.Replace("_params", "");
            File.WriteAllText(pathPackageFile, text1);

            String text2 = File.ReadAllText(txtboxProcess.Text);
            text2 = text2.Replace("***OUTPUT_PATH***", txtboxOutput.Text);

            String pathScriptFile = txtboxProcess.Text.Replace("_params", "");
            File.WriteAllText(pathScriptFile, text2);
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            //exit application when form is closed
            Application.Exit();
        }

        private void btnEditServer_Click(object sender, EventArgs e)
        {
            if(txtboxServer.Enabled)
            {
                btnEditServer.Text = "Edit";
                txtboxServer.Enabled = false;
            }
            else
            {
                btnEditServer.Text = "Save";
                txtboxServer.Enabled = true;
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("C:\\FilePro\\input_template.xls");
        }
    }
}


/*
    String[] sxFiles = Directory.GetFiles(sPath, "output_?.csv");

    // prompt if overwrite
    if (sxFiles.Length > 0)
    {
        DialogResult dialogResult = MessageBox.Show("The folder you specified contains files that would be overwritten.\nWould you like to proceed anyway?", "Warning", MessageBoxButtons.YesNo);
        if (dialogResult == DialogResult.Yes)
        {
*/
