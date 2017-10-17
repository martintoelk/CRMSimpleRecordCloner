namespace martintmg.MSDYN.Tools.SimpleRecordCloner
{
    partial class PluginControl
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(PluginControl));
            this.btnChooseTarget = new System.Windows.Forms.Button();
            this.lstTargetEnvironments = new System.Windows.Forms.ListView();
            this.lblSource = new System.Windows.Forms.Label();
            this.lblRecordURL = new System.Windows.Forms.Label();
            this.txtRecordURL = new System.Windows.Forms.TextBox();
            this.chkIgnoreOwnerAndModifiedBy = new System.Windows.Forms.CheckBox();
            this.chkIgnoreAllLookups = new System.Windows.Forms.CheckBox();
            this.btnCloneRecord = new System.Windows.Forms.Button();
            this.btnAddToRecordList = new System.Windows.Forms.Button();
            this.lblRecordsToProcess = new System.Windows.Forms.Label();
            this.chkVerifyLookups = new System.Windows.Forms.CheckBox();
            this.lstRecordsToProcess = new System.Windows.Forms.ListBox();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.btnClose = new System.Windows.Forms.ToolStripButton();
            this.chkIgnoreStateCode = new System.Windows.Forms.CheckBox();
            this.chkRemoveNonExistingAttributes = new System.Windows.Forms.CheckBox();
            this.chkUpsertRecords = new System.Windows.Forms.CheckBox();
            this.chkRemoveNotFoundLookups = new System.Windows.Forms.CheckBox();
            this.toolStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnChooseTarget
            // 
            this.btnChooseTarget.Location = new System.Drawing.Point(20, 286);
            this.btnChooseTarget.Name = "btnChooseTarget";
            this.btnChooseTarget.Size = new System.Drawing.Size(89, 23);
            this.btnChooseTarget.TabIndex = 0;
            this.btnChooseTarget.Text = "Choose Target";
            this.btnChooseTarget.UseVisualStyleBackColor = true;
            this.btnChooseTarget.Click += new System.EventHandler(this.btnChooseTarget_Click);
            // 
            // lstTargetEnvironments
            // 
            this.lstTargetEnvironments.Location = new System.Drawing.Point(135, 287);
            this.lstTargetEnvironments.Name = "lstTargetEnvironments";
            this.lstTargetEnvironments.Size = new System.Drawing.Size(293, 22);
            this.lstTargetEnvironments.TabIndex = 1;
            this.lstTargetEnvironments.UseCompatibleStateImageBehavior = false;
            this.lstTargetEnvironments.SelectedIndexChanged += new System.EventHandler(this.lstTargetEnvironments_SelectedIndexChanged);
            this.lstTargetEnvironments.KeyDown += new System.Windows.Forms.KeyEventHandler(this.lstTargetEnvironments_KeyDown);
            // 
            // lblSource
            // 
            this.lblSource.AutoSize = true;
            this.lblSource.Location = new System.Drawing.Point(17, 12);
            this.lblSource.Name = "lblSource";
            this.lblSource.Size = new System.Drawing.Size(41, 13);
            this.lblSource.TabIndex = 2;
            this.lblSource.Text = "Source";
            // 
            // lblRecordURL
            // 
            this.lblRecordURL.AutoSize = true;
            this.lblRecordURL.Location = new System.Drawing.Point(17, 43);
            this.lblRecordURL.Name = "lblRecordURL";
            this.lblRecordURL.Size = new System.Drawing.Size(70, 13);
            this.lblRecordURL.TabIndex = 3;
            this.lblRecordURL.Text = "Record URL:";
            // 
            // txtRecordURL
            // 
            this.txtRecordURL.Location = new System.Drawing.Point(135, 37);
            this.txtRecordURL.Name = "txtRecordURL";
            this.txtRecordURL.Size = new System.Drawing.Size(287, 20);
            this.txtRecordURL.TabIndex = 4;
            // 
            // chkIgnoreOwnerAndModifiedBy
            // 
            this.chkIgnoreOwnerAndModifiedBy.AutoSize = true;
            this.chkIgnoreOwnerAndModifiedBy.Location = new System.Drawing.Point(20, 174);
            this.chkIgnoreOwnerAndModifiedBy.Name = "chkIgnoreOwnerAndModifiedBy";
            this.chkIgnoreOwnerAndModifiedBy.Size = new System.Drawing.Size(169, 17);
            this.chkIgnoreOwnerAndModifiedBy.TabIndex = 5;
            this.chkIgnoreOwnerAndModifiedBy.Text = "Ignore Owner and Modified By";
            this.chkIgnoreOwnerAndModifiedBy.UseVisualStyleBackColor = true;
            this.chkIgnoreOwnerAndModifiedBy.CheckedChanged += new System.EventHandler(this.chkIgnoreOwnerAndModifiedBy_CheckedChanged);
            // 
            // chkIgnoreAllLookups
            // 
            this.chkIgnoreAllLookups.AutoSize = true;
            this.chkIgnoreAllLookups.Location = new System.Drawing.Point(20, 197);
            this.chkIgnoreAllLookups.Name = "chkIgnoreAllLookups";
            this.chkIgnoreAllLookups.Size = new System.Drawing.Size(113, 17);
            this.chkIgnoreAllLookups.TabIndex = 6;
            this.chkIgnoreAllLookups.Text = "Ignore all Lookups";
            this.chkIgnoreAllLookups.UseVisualStyleBackColor = true;
            this.chkIgnoreAllLookups.CheckedChanged += new System.EventHandler(this.chkIgnoreAllLookups_CheckedChanged);
            // 
            // btnCloneRecord
            // 
            this.btnCloneRecord.Location = new System.Drawing.Point(20, 336);
            this.btnCloneRecord.Name = "btnCloneRecord";
            this.btnCloneRecord.Size = new System.Drawing.Size(150, 23);
            this.btnCloneRecord.TabIndex = 7;
            this.btnCloneRecord.Text = "Clone Record To Target";
            this.btnCloneRecord.UseVisualStyleBackColor = true;
            this.btnCloneRecord.Click += new System.EventHandler(this.btnCloneRecord_Click);
            // 
            // btnAddToRecordList
            // 
            this.btnAddToRecordList.Location = new System.Drawing.Point(444, 33);
            this.btnAddToRecordList.Name = "btnAddToRecordList";
            this.btnAddToRecordList.Size = new System.Drawing.Size(75, 23);
            this.btnAddToRecordList.TabIndex = 8;
            this.btnAddToRecordList.Text = "Add To List";
            this.btnAddToRecordList.UseVisualStyleBackColor = true;
            this.btnAddToRecordList.Click += new System.EventHandler(this.btnAddToRecordList_Click);
            // 
            // lblRecordsToProcess
            // 
            this.lblRecordsToProcess.AutoSize = true;
            this.lblRecordsToProcess.Location = new System.Drawing.Point(17, 81);
            this.lblRecordsToProcess.Name = "lblRecordsToProcess";
            this.lblRecordsToProcess.Size = new System.Drawing.Size(103, 13);
            this.lblRecordsToProcess.TabIndex = 10;
            this.lblRecordsToProcess.Text = "Records to Process:";
            // 
            // chkVerifyLookups
            // 
            this.chkVerifyLookups.AutoSize = true;
            this.chkVerifyLookups.Location = new System.Drawing.Point(300, 174);
            this.chkVerifyLookups.Name = "chkVerifyLookups";
            this.chkVerifyLookups.Size = new System.Drawing.Size(203, 17);
            this.chkVerifyLookups.TabIndex = 11;
            this.chkVerifyLookups.Text = "Verify Lookups in Target Environment";
            this.chkVerifyLookups.UseVisualStyleBackColor = true;
            this.chkVerifyLookups.CheckedChanged += new System.EventHandler(this.chkVerifyLookups_CheckedChanged);
            // 
            // lstRecordsToProcess
            // 
            this.lstRecordsToProcess.FormattingEnabled = true;
            this.lstRecordsToProcess.Location = new System.Drawing.Point(135, 73);
            this.lstRecordsToProcess.Name = "lstRecordsToProcess";
            this.lstRecordsToProcess.Size = new System.Drawing.Size(293, 95);
            this.lstRecordsToProcess.TabIndex = 13;
            this.lstRecordsToProcess.KeyDown += new System.Windows.Forms.KeyEventHandler(this.lstRecordsToProcess_KeyDown);
            // 
            // toolStrip1
            // 
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnClose});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(543, 25);
            this.toolStrip1.TabIndex = 14;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // btnClose
            // 
            this.btnClose.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.btnClose.Image = ((System.Drawing.Image)(resources.GetObject("btnClose.Image")));
            this.btnClose.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(23, 22);
            this.btnClose.Text = "Close This Tool";
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // chkIgnoreStateCode
            // 
            this.chkIgnoreStateCode.AutoSize = true;
            this.chkIgnoreStateCode.Location = new System.Drawing.Point(300, 197);
            this.chkIgnoreStateCode.Name = "chkIgnoreStateCode";
            this.chkIgnoreStateCode.Size = new System.Drawing.Size(141, 17);
            this.chkIgnoreStateCode.TabIndex = 15;
            this.chkIgnoreStateCode.Text = "IgnoreStateStatusCodes";
            this.chkIgnoreStateCode.UseVisualStyleBackColor = true;
            // 
            // chkRemoveNonExistingAttributes
            // 
            this.chkRemoveNonExistingAttributes.AutoSize = true;
            this.chkRemoveNonExistingAttributes.Location = new System.Drawing.Point(20, 220);
            this.chkRemoveNonExistingAttributes.Name = "chkRemoveNonExistingAttributes";
            this.chkRemoveNonExistingAttributes.Size = new System.Drawing.Size(257, 17);
            this.chkRemoveNonExistingAttributes.TabIndex = 16;
            this.chkRemoveNonExistingAttributes.Text = "Remove missing Attributes in Traget Environment";
            this.chkRemoveNonExistingAttributes.UseVisualStyleBackColor = true;
            // 
            // chkUpsertRecords
            // 
            this.chkUpsertRecords.AutoSize = true;
            this.chkUpsertRecords.Location = new System.Drawing.Point(300, 220);
            this.chkUpsertRecords.Name = "chkUpsertRecords";
            this.chkUpsertRecords.Size = new System.Drawing.Size(97, 17);
            this.chkUpsertRecords.TabIndex = 17;
            this.chkUpsertRecords.Text = "Upser Records";
            this.chkUpsertRecords.UseVisualStyleBackColor = true;
            // 
            // chkRemoveNotFoundLookups
            // 
            this.chkRemoveNotFoundLookups.AutoSize = true;
            this.chkRemoveNotFoundLookups.Location = new System.Drawing.Point(20, 244);
            this.chkRemoveNotFoundLookups.Name = "chkRemoveNotFoundLookups";
            this.chkRemoveNotFoundLookups.Size = new System.Drawing.Size(315, 17);
            this.chkRemoveNotFoundLookups.TabIndex = 18;
            this.chkRemoveNotFoundLookups.Text = "Remove Lookups which are not present in the Target System";
            this.chkRemoveNotFoundLookups.UseVisualStyleBackColor = true;
            // 
            // PluginControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.chkRemoveNotFoundLookups);
            this.Controls.Add(this.chkUpsertRecords);
            this.Controls.Add(this.chkRemoveNonExistingAttributes);
            this.Controls.Add(this.chkIgnoreStateCode);
            this.Controls.Add(this.toolStrip1);
            this.Controls.Add(this.lstRecordsToProcess);
            this.Controls.Add(this.chkVerifyLookups);
            this.Controls.Add(this.lblRecordsToProcess);
            this.Controls.Add(this.btnAddToRecordList);
            this.Controls.Add(this.btnCloneRecord);
            this.Controls.Add(this.chkIgnoreAllLookups);
            this.Controls.Add(this.chkIgnoreOwnerAndModifiedBy);
            this.Controls.Add(this.txtRecordURL);
            this.Controls.Add(this.lblRecordURL);
            this.Controls.Add(this.lblSource);
            this.Controls.Add(this.lstTargetEnvironments);
            this.Controls.Add(this.btnChooseTarget);
            this.Name = "PluginControl";
            this.Size = new System.Drawing.Size(543, 412);
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnChooseTarget;
        private System.Windows.Forms.ListView lstTargetEnvironments;
        private System.Windows.Forms.Label lblSource;
        private System.Windows.Forms.Label lblRecordURL;
        private System.Windows.Forms.TextBox txtRecordURL;
        private System.Windows.Forms.CheckBox chkIgnoreOwnerAndModifiedBy;
        private System.Windows.Forms.CheckBox chkIgnoreAllLookups;
        private System.Windows.Forms.Button btnCloneRecord;
        private System.Windows.Forms.Button btnAddToRecordList;
        private System.Windows.Forms.Label lblRecordsToProcess;
        private System.Windows.Forms.CheckBox chkVerifyLookups;
        private System.Windows.Forms.ListBox lstRecordsToProcess;
        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripButton btnClose;
        private System.Windows.Forms.CheckBox chkIgnoreStateCode;
        private System.Windows.Forms.CheckBox chkRemoveNonExistingAttributes;
        private System.Windows.Forms.CheckBox chkUpsertRecords;
        private System.Windows.Forms.CheckBox chkRemoveNotFoundLookups;
    }
}
