using Microsoft.Crm.Sdk.Messages;
using System;
using System.Windows.Forms;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Interfaces;
using XrmToolBox.Extensibility.Args;
using Microsoft.Xrm.Sdk;
using McTools.Xrm.Connection;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Messages;

namespace martintmg.MSDYN.Tools.SimpleRecordCloner
{
    public partial class PluginControl : UserControl, IXrmToolBoxPluginControl, IGitHubPlugin, IStatusBarMessenger, IHelpPlugin
    {
        private ConnectionDetail detail;

        private IOrganizationService service;
        private Dictionary<string, IOrganizationService> targetServices = new Dictionary<string, IOrganizationService>();
        private KeyValuePair<string, IOrganizationService> lastTargetService;

        public string RepositoryName
        {
            get
            {
                return "CRMSimpleRecordCloner";
            }
        }

        public string UserName
        {
            get
            {
                return "martintmg";
            }
        }

        public string HelpUrl
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public IOrganizationService Service
        {
            get
            {
                throw new NotImplementedException();
            }
        }
        #region XrmToolbox
        public event EventHandler OnRequestConnection;
        #endregion XrmToolbox


        private void SetConnectionLabel(ConnectionDetail detail, string serviceType)
        {
            switch (serviceType)
            {
                case "Source":
                    lblSource.Text = detail.ConnectionName;
                    lblSource.ForeColor = Color.Green;
                    break;

                    //case "Target":
                    //    lblTarget.Text = detail.ConnectionName;
                    //    lblTarget.ForeColor = Color.Green;
                    //    break;
            }
        }
       
        public  void UpdateConnection(IOrganizationService newService, ConnectionDetail detail, string actionName = "", object parameter = null)
        {
            this.detail = detail;
            if (actionName == "TargetOrganization")
            {
                targetServices.Add(detail.ConnectionName, newService);
                lastTargetService = new KeyValuePair<string, IOrganizationService>(detail.ConnectionName, newService);
                lstTargetEnvironments.Items.Add(new ListViewItem { Text = detail.ConnectionName });
            }
            else
            {
                service = newService;

                SetConnectionLabel(detail, "Source");

            }
        }

        #region Base tool implementation

        public PluginControl()
        {
            InitializeComponent();
        }

        public event EventHandler<StatusBarMessageEventArgs> SendMessageToStatusBar;
        public event EventHandler OnCloseTool;

       
        #endregion Base tool implementation

        private void btnChooseTarget_Click(object sender, EventArgs e)
        {
            if (OnRequestConnection != null)
            {
                var args = new RequestConnectionEventArgs { ActionName = "TargetOrganization", Control = this };
                OnRequestConnection(this, args);
            }
        }
        private void chkIgnoreAllLookups_CheckedChanged(object sender, EventArgs e)
        {
            if (((CheckBox)sender).Checked == true)
            {
                chkIgnoreOwnerAndModifiedBy.Checked = true;
                chkVerifyLookups.Checked = false;
            }
        }

        private void chkIgnoreOwnerAndModifiedBy_CheckedChanged(object sender, EventArgs e)
        {
            if (((CheckBox)sender).Checked == false)
                chkIgnoreAllLookups.Checked = false;
        }

        private void lstTargetEnvironments_KeyDown(object sender, KeyEventArgs e)
        {
            if (lstTargetEnvironments.SelectedItems.Count > 0 && e.KeyCode == Keys.Delete)
            {
                foreach (ListViewItem item in lstTargetEnvironments.Items)
                {
                    if (item.Selected)
                        lstTargetEnvironments.Items.Remove(item);
                }

                ReorderTargetServices();
            }
        }

        private void ReorderTargetServices()
        {
            var oldTargetServices = targetServices;
            targetServices = new Dictionary<string, IOrganizationService>();

            foreach (ListViewItem item in lstTargetEnvironments.Items)
            {
                var serviceToAdd = oldTargetServices.First(x => x.Key == item.Text);
                targetServices.Add(serviceToAdd.Key, serviceToAdd.Value);
            }
        }

        private void btnCloneRecord_Click(object sender, EventArgs e)
        {
            txtRecordURL.Enabled = false;
            btnAddToRecordList.Enabled = false;
            btnChooseTarget.Enabled = false;

            var recordToProcess = lstRecordsToProcess.Items;

            foreach (var targetService in targetServices)
            {
                lastTargetService = targetService;
            }

            foreach (var record in recordToProcess)
            {
                string logicalName;
                Guid id;
                GetParameterFromURL(record.ToString(), out logicalName, out id);

                var entity = service.Retrieve(logicalName, id, new ColumnSet(true));

                if (chkIgnoreAllLookups.Checked)
                {
                    RemoveERefs(entity);
                }

                if (chkIgnoreOwnerAndModifiedBy.Checked)
                {
                    RemoveOwnerAndLastmodifed(entity);
                }

                if (chkVerifyLookups.Checked)
                {
                    CheckLookups(entity);
                }

                lastTargetService.Value.Create(entity);
            }

            txtRecordURL.Enabled = true;
            btnAddToRecordList.Enabled = true;
            btnChooseTarget.Enabled = true;
        }

        private void GetParameterFromURL(string RecordUrl, out string LogicalName, out Guid RecordId)
        {
            Uri url;
            Uri.TryCreate(RecordUrl, UriKind.Absolute, out url);
            int typeCode = -1;
            var query = url.Query.Replace("?", "");
            RecordId = Guid.Empty;

            var parameters = query.Split('&').ToList();

            foreach(var parameter in parameters)
            {
                var nameAndValue = parameter.Split('=');

                if (nameAndValue[0] == "etc")
                {
                    typeCode = int.Parse(nameAndValue[1]);
                }

                if (nameAndValue[0] == "id")
                {
                    RecordId = new Guid(nameAndValue[1].Replace("%7b", "").Replace("%7d", ""));
                }
            }
            if (typeCode == -1)
            {
                throw new ArgumentException("EntityTypeCode can't be pasred from URL");
            }

            if (RecordId == Guid.Empty)
            {
                throw new ArgumentException("Entity Id can't be pasred from URL");
            }

            var metaDataRequest = new RetrieveAllEntitiesRequest();
            metaDataRequest.RetrieveAsIfPublished = true;

            var metadata = (RetrieveAllEntitiesResponse)service.Execute(metaDataRequest);

            var targetEntity = metadata.EntityMetadata.FirstOrDefault(entity => entity.ObjectTypeCode == typeCode);

            LogicalName = targetEntity.LogicalName;
        }

        private void CheckLookups(Entity Entity)
        {
            var allERefsAttributes = Entity.Attributes.Where(attribute => attribute.Value.GetType().Name == "EntityReference").ToList();

            var errorLookups = new List<EntityReference>();

            foreach (var eRef in allERefsAttributes)
            {
                try
                {
                    var targetEntity = lastTargetService.Value.Retrieve(((EntityReference)eRef.Value).LogicalName, ((EntityReference)eRef.Value).Id, new ColumnSet());
                }
                catch
                {
                    errorLookups.Add((EntityReference)eRef.Value);
                }
            }

            if (errorLookups.Any())
            {
                var message = "The following Lookups are not present in the Target Enviornment" + Environment.NewLine;

                foreach (var error in errorLookups)
                {
                    message += string.Format("LogicalName {0} Id {1} does not exist {2}", error.LogicalName, error.Id, Environment.NewLine);
                }

                MessageBox.Show(message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                return;
            }
        }

        private static void RemoveOwnerAndLastmodifed(Entity entity)
        {
            if (entity.Contains("ownerid"))
            {
                entity.Attributes.Remove("ownerid");
            }
            if (entity.Contains("modifiedby"))
            {
                entity.Attributes.Remove("modifiedby");
            }
        }

        private static void RemoveERefs(Entity entity)
        {
            var allAttributesWithOutERefs = entity.Attributes.Where(attribute => attribute.Value.GetType().Name != "EntityReference").ToList();

            entity.Attributes = new AttributeCollection();
            entity.Attributes.AddRange(allAttributesWithOutERefs);
        }

        private void chkVerifyLookups_CheckedChanged(object sender, EventArgs e)
        {
            if (((CheckBox)sender).Checked == true)
            {
                chkIgnoreAllLookups.Checked = false;
            }
        }

        private void btnAddToRecordList_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(txtRecordURL.Text))
            {
                lstRecordsToProcess.Items.Add(txtRecordURL.Text);
            }
        }

        private void lstTargetEnvironments_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        public void ClosingPlugin(PluginCloseInfo info)
        {
            info.Cancel = MessageBox.Show(@"Are you sure you want to close this tab?", @"Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes;
        }
    }
}