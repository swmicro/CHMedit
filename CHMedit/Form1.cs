using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Xml;
using System.Net;
using System.Diagnostics;
using System.Security.AccessControl;
using System.Security.Permissions;
using System.Security.Principal;

namespace CHMedit
{
    public partial class Form1 : Form
    {
        string fileName = null; 
        string fullFileName = null;
        string projectFileName = null;
        string editorFileName = null;
        string outputDir = null; 
        string hhcFilePath = null;
        string hhFilePath = null;

        System.Diagnostics.Process myProc = new System.Diagnostics.Process();

        List<string> _items = new List<string>();
        
        public Form1()
        {
            InitializeComponent();

            if (File.Exists(@"C:\Program Files (x86)\Notepad++\notepad++.exe"))
                editorFileName = @"C:\Program Files (x86)\Notepad++\notepad++.exe";
            else if (File.Exists(@"C:\Program Files\Notepad++\notepad++.exe"))
                editorFileName = @"C:\Program Files\Notepad++\notepad++.exe";
            else
                editorFileName = "notepad";

            if (File.Exists(@"C:\Program Files (x86)\HTML Help Workshop\hhc.exe"))
            {
                hhcFilePath = @"C:\Program Files (x86)\HTML Help Workshop\hhc.exe";
                hhFilePath = @"C:\Program Files (x86)\HTML Help Workshop\hh.exe";
            }
            else if (File.Exists(@"C:\Program Files\HTML Help Workshop\hhc.exe"))
            {
                hhcFilePath = @"C:\Program Files\HTML Help Workshop\hhc.exe";
                hhFilePath = @"C:\Program Files\HTML Help Workshop\hh.exe";
            } 
            else
                toolStripStatusLabel1.Text = "Error: HTML Help Workshop is not found!";

            if (hhcFilePath != null)
                SetAcl(Path.GetDirectoryName(hhcFilePath));

            if (hhcFilePath != null)
                SetAcl(Path.GetDirectoryName(hhFilePath));
        }

        private bool openCHMFile()
        {
            bool result = false;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.Filter = "CHM files (*.chm) | *.chm";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fileName = openFileDialog1.FileName.ToString();
                fullFileName = Path.GetFullPath(fileName);
                SetAcl(fullFileName);
                projectFileName = Path.GetFileName(fileName);
                projectFileName = projectFileName.Replace(".chm", "");

                outputDir = Path.GetDirectoryName(fileName) + "\\decompiledChm";
                if (Directory.Exists(outputDir) == true)
                    Directory.Delete(Path.GetFullPath(outputDir), true);
                Directory.CreateDirectory(outputDir);
                SetAccessRule(outputDir);
                SetAcl(outputDir);

                result = true;
            }

            return result;
         }

        public void RunApp(string FileName, string arg)
        {
            SecurityPermission SP = new SecurityPermission(SecurityPermissionFlag.AllFlags);
            SP.Assert();

            //ProcessStartInfo process = new ProcessStartInfo("cmd", "/c " + FileName )); 
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            process.StartInfo.WindowStyle = ProcessWindowStyle.Normal;
            //process.StartInfo.UseShellExecute = false;
            //process.StartInfo.RedirectStandardOutput = true;
            //process.StartInfo.RedirectStandardError = true;
            //process.StartInfo.CreateNoWindow = true;
            process.StartInfo.FileName = FileName;
            process.StartInfo.Arguments = arg;
            process.StartInfo.WorkingDirectory = System.IO.Path.GetDirectoryName(FileName);

            //Vista or higher check
            if (System.Environment.OSVersion.Version.Major >= 6)
                process.StartInfo.Verb = "runas";

            try
            {
                process.Start();
                process.WaitForExit();
                process.Close();
            }
            catch (InvalidOperationException)
            {
                //e.ExceptionObject.ToString();
            }
        }

        public static void SetAccessRule(string directory)
        {
            DirectorySecurity sec = Directory.GetAccessControl(directory);
            SecurityIdentifier everyone = new SecurityIdentifier( WellKnownSidType.WorldSid, null);
            sec.AddAccessRule(
                new FileSystemAccessRule(
                    everyone, 
                    FileSystemRights.FullControl, 
                    InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit, 
                    PropagationFlags.None, 
                    AccessControlType.Allow)
                    );
            Directory.SetAccessControl(directory, sec);
        }

        static bool SetAcl(string directory)
        {
            FileSystemRights Rights = (FileSystemRights)0;
            Rights = FileSystemRights.FullControl;

            // *** Add Access Rule to the actual directory itself
            FileSystemAccessRule AccessRule = new FileSystemAccessRule(
                @"BUILTIN\Users", 
                Rights, 
                InheritanceFlags.None,
                PropagationFlags.NoPropagateInherit,
                AccessControlType.Allow
                );

            DirectoryInfo Info = new DirectoryInfo(directory);
            DirectorySecurity Security = Info.GetAccessControl(AccessControlSections.Access);

            bool Result = false;
            Security.ModifyAccessRule(AccessControlModification.Set, AccessRule, out Result);

            if (!Result)
                return false;

            // *** Always allow objects to inherit on a directory
            InheritanceFlags iFlags = InheritanceFlags.ObjectInherit;
            iFlags = InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit;

            // *** Add Access rule for the inheritance
            AccessRule = new FileSystemAccessRule(
                        @"BUILTIN\Users", 
                        Rights,
                        iFlags,
                        PropagationFlags.InheritOnly,
                        AccessControlType.Allow
                        );
            Result = false;
            Security.ModifyAccessRule(AccessControlModification.Add, AccessRule, out Result);

            if (!Result)
                return false;

            Info.SetAccessControl(Security);

            return true;
        }

        private void openCHM()
        {
            string str;

            if (openCHMFile() == false)
                return;

            // Decompiling a CHM file
            ///System.Diagnostics.Process.Start(hhFilePath, str);
            //hh.exe -decompile <target-folder-for-decompiled-content> <source-chm-file>
            str = "-decompile " + outputDir + " " + fileName;
            RunApp(hhFilePath, str);
            createHHP(null);
            webBrowser1.Navigate(outputDir);

            str = outputDir + "\\Release Notes.htm";
            if (File.Exists(str))
                Process.Start(editorFileName, str);
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            openCHM();
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            compileCHM();
        }

        private void openCHMToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openCHM();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show(
                "To use this app you must have HTML Help Workshop \n" +
                " http://go.microsoft.com/fwlink/?LinkId=14188  \n" +
//                "and an optional is notepad++ http://notepad-plus-plus.org/download \n" +
                "\n" +
                "How it works: \n" +
                "1. Press Open CHM file \n" +
                "2. When decompiling done you can edit decompiled HTML files \n" +
                "3. Press Compile CHM file.  \n" +
                "4. Press Show CHM file to see what you've done.  \n\n" +
                "support email: vyacheslava@ami.com (swmicro@gmail.com) \n"
            );
        }

        private void selectEditorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog2 = new OpenFileDialog();

            if (Directory.Exists("C:\\Program Files\\"))
                openFileDialog2.InitialDirectory = "C:\\Program Files\\";
            else if (Directory.Exists("C:\\Program Files (x86)\\"))
                openFileDialog2.InitialDirectory = "C:\\Program Files (x86)\\";
            else
                openFileDialog2.InitialDirectory = "C:\\";

            openFileDialog2.Filter = "EXE files (*.exe) | *.exe";
            openFileDialog2.FilterIndex = 2;
            openFileDialog2.RestoreDirectory = true;

            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                string str = openFileDialog2.FileName.ToString();
                editorFileName = Path.GetFullPath(str);
            }
        }

        private void showCHMToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (projectFileName == null)
                MessageBox.Show("Compile a CHM file first");

            Help.ShowHelp(this, outputDir + "\\" + projectFileName + ".chm");
        }

        private void compileCHMToolStripMenuItem_Click(object sender, EventArgs e)
        {
            compileCHM();
        }

        private void createHHP(string hhpFile)
        {
            if (hhpFile == null)
                hhpFile = outputDir + "\\" + projectFileName + ".hhp";

            if (File.Exists(hhpFile))
                File.Delete(hhpFile);

            FileStream fs = new FileStream(hhpFile, FileMode.Create);
            StreamWriter w = new StreamWriter(fs, Encoding.UTF8);
            w.WriteLine("[OPTIONS]");
            w.WriteLine("Binary TOC=Yes");
            w.WriteLine("Compatibility=1.1 or later");
            //w.WriteLine("Compiled file=..\ .chm");
            w.WriteLine("Contents file=Table of Contents.hhc");
            w.WriteLine("Default topic=Main Page.htm");
            w.WriteLine("Display compile progress=No");
            w.WriteLine("Full-text search=Yes");
            w.WriteLine("Index file=Index.hhk");
            w.WriteLine("Language=0x409 English (United States)");
            w.WriteLine("[FILES]");

            string[] files = Directory.GetFiles(outputDir, "*.htm");
            foreach (string f in files)
            {
                w.WriteLine(Path.GetFileName(f));
            }
            w.Close();
        }

        private void compileCHM()
        {
            string fileHHP = outputDir + "\\" + projectFileName + ".hhp";

            if (fullFileName == null)
            {
                MessageBox.Show("Open a CHM file first");
                return;
            }

            createHHP(fileHHP);
            RunApp(hhcFilePath, fileHHP);
            webBrowser1.Navigate(outputDir);
        }

        private void saveCHMToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFile = new SaveFileDialog();

            saveFile.Filter = "CHM files (*.chm) | *.chm";
            saveFile.FilterIndex = 2;
            saveFile.RestoreDirectory = true;
            saveFile.FileName = outputDir + "\\" + projectFileName + ".chm";

            if (saveFile.ShowDialog() == DialogResult.OK)
            {
                toolStripStatusLabel1.Text = "Saved";
            }
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            if (projectFileName == null)
                MessageBox.Show("Compile a CHM file first");

            Help.ShowHelp(this, outputDir + "\\" + projectFileName + ".chm");
        }

        private void setHTMLFolderToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog openFolder = new FolderBrowserDialog();

            if (openFolder.ShowDialog() == DialogResult.OK)
            {
                string str = openFolder.SelectedPath;
                hhcFilePath = str + "\\hhc.exe"; 
            }
        }

    }
}
