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
            this.btnChooseTarget = new System.Windows.Forms.Button();
            this.lstTargetEnvironments = new System.Windows.Forms.ListView();
            this.lblSource = new System.Windows.Forms.Label();
            this.lblRecordURL = new System.Windows.Forms.Label();
            this.txtRecordURL = new System.Windows.Forms.TextBox();
            this.chkIgnoreOwnerAndModifiedBy = new System.Windows.Forms.CheckBox();
            this.chkIgnoreAllLookups = new System.Windows.Forms.CheckBox();
            this.btnCloneRecord = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnChooseTarget
            // 
            this.btnChooseTarget.Location = new System.Drawing.Point(20, 108);
            this.btnChooseTarget.Name = "btnChooseTarget";
            this.btnChooseTarget.Size = new System.Drawing.Size(89, 23);
            this.btnChooseTarget.TabIndex = 0;
            this.btnChooseTarget.Text = "Choose Target";
            this.btnChooseTarget.UseVisualStyleBackColor = true;
            this.btnChooseTarget.Click += new System.EventHandler(this.btnChooseTarget_Click);
            // 
            // lstTargetEnvironments
            // 
            this.lstTargetEnvironments.Location = new System.Drawing.Point(135, 109);
            this.lstTargetEnvironments.Name = "lstTargetEnvironments";
            this.lstTargetEnvironments.Size = new System.Drawing.Size(293, 22);
            this.lstTargetEnvironments.TabIndex = 1;
            this.lstTargetEnvironments.UseCompatibleStateImageBehavior = false;
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
            this.lblRecordURL.Location = new System.Drawing.Point(17, 44);
            this.lblRecordURL.Name = "lblRecordURL";
            this.lblRecordURL.Size = new System.Drawing.Size(67, 13);
            this.lblRecordURL.TabIndex = 3;
            this.lblRecordURL.Text = "Record URL";
            // 
            // txtRecordURL
            // 
            this.txtRecordURL.Location = new System.Drawing.Point(102, 37);
            this.txtRecordURL.Name = "txtRecordURL";
            this.txtRecordURL.Size = new System.Drawing.Size(298, 20);
            this.txtRecordURL.TabIndex = 4;
            // 
            // chkIgnoreOwnerAndModifiedBy
            // 
            this.chkIgnoreOwnerAndModifiedBy.AutoSize = true;
            this.chkIgnoreOwnerAndModifiedBy.Location = new System.Drawing.Point(20, 63);
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
            this.chkIgnoreAllLookups.Location = new System.Drawing.Point(20, 86);
            this.chkIgnoreAllLookups.Name = "chkIgnoreAllLookups";
            this.chkIgnoreAllLookups.Size = new System.Drawing.Size(113, 17);
            this.chkIgnoreAllLookups.TabIndex = 6;
            this.chkIgnoreAllLookups.Text = "Ignore all Lookups";
            this.chkIgnoreAllLookups.UseVisualStyleBackColor = true;
            this.chkIgnoreAllLookups.CheckedChanged += new System.EventHandler(this.chkIgnoreAllLookups_CheckedChanged);
            // 
            // btnCloneRecord
            // 
            this.btnCloneRecord.Location = new System.Drawing.Point(20, 165);
            this.btnCloneRecord.Name = "btnCloneRecord";
            this.btnCloneRecord.Size = new System.Drawing.Size(150, 23);
            this.btnCloneRecord.TabIndex = 7;
            this.btnCloneRecord.Text = "Clone Record To Target";
            this.btnCloneRecord.UseVisualStyleBackColor = true;
            this.btnCloneRecord.Click += new System.EventHandler(this.btnCloneRecord_Click);
            // 
            // PluginControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
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
    }
}
