// DataExtractor is an ArcGIS add-in used to extract biodiversity
// information from SQL Server based on existing boundaries.
//
// Copyright © 2017 SxBRC, 2017-2019 TVERC, 2020 Andy Foy Consulting
//
// This file is part of DataExtractor.
//
// DataExtractor is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// DataExtractor is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with DataExtractor.  If not, see <http://www.gnu.org/licenses/>.

namespace DataExtractor
{
    partial class frmDataExtractor
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmDataExtractor));
            this.label1 = new System.Windows.Forms.Label();
            this.lstActivePartners = new System.Windows.Forms.ListBox();
            this.chkZip = new System.Windows.Forms.CheckBox();
            this.chkApplyExclusion = new System.Windows.Forms.CheckBox();
            this.chkClearLog = new System.Windows.Forms.CheckBox();
            this.label2 = new System.Windows.Forms.Label();
            this.lstTables = new System.Windows.Forms.ListBox();
            this.label3 = new System.Windows.Forms.Label();
            this.lstLayers = new System.Windows.Forms.ListBox();
            this.cmbSelectionType = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.lblStatus = new System.Windows.Forms.Label();
            this.chkUseCentroids = new System.Windows.Forms.CheckBox();
            this.chkUploadToServer = new System.Windows.Forms.CheckBox();
            this.btnAbout = new System.Windows.Forms.Button();
            this.toolTip = new System.Windows.Forms.ToolTip(this.components);
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(82, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Active Partners:";
            // 
            // lstActivePartners
            // 
            this.lstActivePartners.FormattingEnabled = true;
            this.lstActivePartners.Location = new System.Drawing.Point(16, 31);
            this.lstActivePartners.Name = "lstActivePartners";
            this.lstActivePartners.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.lstActivePartners.Size = new System.Drawing.Size(200, 290);
            this.lstActivePartners.TabIndex = 1;
            this.lstActivePartners.DoubleClick += new System.EventHandler(this.lstActivePartners_DoubleClick);
            // 
            // chkZip
            // 
            this.chkZip.AutoSize = true;
            this.chkZip.Location = new System.Drawing.Point(18, 332);
            this.chkZip.Name = "chkZip";
            this.chkZip.Size = new System.Drawing.Size(103, 17);
            this.chkZip.TabIndex = 2;
            this.chkZip.Text = "Zip extract files?";
            this.chkZip.UseVisualStyleBackColor = true;
            this.chkZip.Visible = false;
            // 
            // chkApplyExclusion
            // 
            this.chkApplyExclusion.AutoSize = true;
            this.chkApplyExclusion.Location = new System.Drawing.Point(18, 356);
            this.chkApplyExclusion.Name = "chkApplyExclusion";
            this.chkApplyExclusion.Size = new System.Drawing.Size(139, 17);
            this.chkApplyExclusion.TabIndex = 3;
            this.chkApplyExclusion.Text = "Apply exclusion clause?";
            this.chkApplyExclusion.UseVisualStyleBackColor = true;
            // 
            // chkClearLog
            // 
            this.chkClearLog.AutoSize = true;
            this.chkClearLog.Location = new System.Drawing.Point(18, 379);
            this.chkClearLog.Name = "chkClearLog";
            this.chkClearLog.Size = new System.Drawing.Size(89, 17);
            this.chkClearLog.TabIndex = 4;
            this.chkClearLog.Text = "Clear log file?";
            this.chkClearLog.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(228, 13);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(66, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "SQL Tables:";
            // 
            // lstTables
            // 
            this.lstTables.FormattingEnabled = true;
            this.lstTables.Location = new System.Drawing.Point(231, 31);
            this.lstTables.Name = "lstTables";
            this.lstTables.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.lstTables.Size = new System.Drawing.Size(200, 290);
            this.lstTables.TabIndex = 6;
            this.lstTables.DoubleClick += new System.EventHandler(this.lstTables_DoubleClick);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(443, 13);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(62, 13);
            this.label3.TabIndex = 9;
            this.label3.Text = "GIS Layers:";
            // 
            // lstLayers
            // 
            this.lstLayers.FormattingEnabled = true;
            this.lstLayers.Location = new System.Drawing.Point(446, 31);
            this.lstLayers.Name = "lstLayers";
            this.lstLayers.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.lstLayers.Size = new System.Drawing.Size(200, 290);
            this.lstLayers.TabIndex = 10;
            this.lstLayers.DoubleClick += new System.EventHandler(this.lstLayers_DoubleClick);
            // 
            // cmbSelectionType
            // 
            this.cmbSelectionType.FormattingEnabled = true;
            this.cmbSelectionType.Location = new System.Drawing.Point(448, 352);
            this.cmbSelectionType.Name = "cmbSelectionType";
            this.cmbSelectionType.Size = new System.Drawing.Size(200, 21);
            this.cmbSelectionType.TabIndex = 12;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(445, 332);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(81, 13);
            this.label4.TabIndex = 11;
            this.label4.Text = "Selection Type:";
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(531, 396);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(55, 23);
            this.btnCancel.TabIndex = 13;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(592, 396);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(56, 23);
            this.btnOK.TabIndex = 14;
            this.btnOK.Text = "Ok";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new System.Drawing.Point(15, 406);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(43, 13);
            this.lblStatus.TabIndex = 15;
            this.lblStatus.Text = "Status: ";
            this.lblStatus.Visible = false;
            // 
            // chkUseCentroids
            // 
            this.chkUseCentroids.AutoSize = true;
            this.chkUseCentroids.Location = new System.Drawing.Point(233, 332);
            this.chkUseCentroids.Name = "chkUseCentroids";
            this.chkUseCentroids.Size = new System.Drawing.Size(167, 17);
            this.chkUseCentroids.TabIndex = 7;
            this.chkUseCentroids.Text = "Select polygons by centroids?";
            this.chkUseCentroids.UseVisualStyleBackColor = true;
            // 
            // chkUploadToServer
            // 
            this.chkUploadToServer.AutoSize = true;
            this.chkUploadToServer.Location = new System.Drawing.Point(233, 356);
            this.chkUploadToServer.Name = "chkUploadToServer";
            this.chkUploadToServer.Size = new System.Drawing.Size(172, 17);
            this.chkUploadToServer.TabIndex = 8;
            this.chkUploadToServer.Text = "Upload partner table to server?";
            this.chkUploadToServer.UseVisualStyleBackColor = true;
            // 
            // btnAbout
            // 
            this.btnAbout.Location = new System.Drawing.Point(625, 5);
            this.btnAbout.Margin = new System.Windows.Forms.Padding(0);
            this.btnAbout.Name = "btnAbout";
            this.btnAbout.Size = new System.Drawing.Size(21, 21);
            this.btnAbout.TabIndex = 16;
            this.btnAbout.TabStop = false;
            this.btnAbout.Text = "?";
            this.toolTip.SetToolTip(this.btnAbout, "About DataExtractor");
            this.btnAbout.UseVisualStyleBackColor = true;
            this.btnAbout.Click += new System.EventHandler(this.btnAbout_Click);
            // 
            // frmDataExtractor
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(661, 431);
            this.Controls.Add(this.btnAbout);
            this.Controls.Add(this.chkUploadToServer);
            this.Controls.Add(this.chkUseCentroids);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.cmbSelectionType);
            this.Controls.Add(this.lstLayers);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.lstTables);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.chkClearLog);
            this.Controls.Add(this.chkApplyExclusion);
            this.Controls.Add(this.chkZip);
            this.Controls.Add(this.lstActivePartners);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmDataExtractor";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Data Extractor 11.0.0";
            this.Load += new System.EventHandler(this.frmDataExtractor_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListBox lstActivePartners;
        private System.Windows.Forms.CheckBox chkZip;
        private System.Windows.Forms.CheckBox chkApplyExclusion;
        private System.Windows.Forms.CheckBox chkClearLog;
        private System.Windows.Forms.CheckBox chkUseCentroids;
        private System.Windows.Forms.CheckBox chkUploadToServer;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ListBox lstTables;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ListBox lstLayers;
        private System.Windows.Forms.ComboBox cmbSelectionType;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Button btnAbout;
        private System.Windows.Forms.ToolTip toolTip;
    }
}