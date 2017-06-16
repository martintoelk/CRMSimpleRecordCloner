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

namespace martintmg.MSDYN.Tools.SimpleRecordCloner
{
    public partial class PluginControl : PluginControlBase, IGitHubPlugin, IStatusBarMessenger, IHelpPlugin
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

        public override void UpdateConnection(IOrganizationService newService, ConnectionDetail detail, string actionName = "", object parameter = null)
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

        public void ProcessWhoAmI()
        {
            WorkAsync(new WorkAsyncInfo
            {
                Message = "Retrieving your user id...",
                Work = (w, e) =>
                {
                    var request = new WhoAmIRequest();
                    var response = (WhoAmIResponse)Service.Execute(request);

                    e.Result = response.UserId;
                },
                ProgressChanged = e =>
                {
                    // If progress has to be notified to user, use the following method:
                    SetWorkingMessage("Message to display");
                },
                PostWorkCallBack = e =>
                {
                    MessageBox.Show(string.Format("You are {0}", (Guid)e.Result));
                },
                AsyncArgument = null,
                IsCancelable = true,
                MessageWidth = 340,
                MessageHeight = 150
            });
        }

        private void BtnCloseClick(object sender, EventArgs e)
        {
            CloseTool(); // PluginBaseControl method that notifies the XrmToolBox that the user wants to close the plugin
            // Override the ClosingPlugin method to allow for any plugin specific closing logic to be performed (saving configs, canceling close, etc...)
        }

        private void BtnWhoAmIClick(object sender, EventArgs e)
        {
            ExecuteMethod(ProcessWhoAmI); // ExecuteMethod ensures that the user has connected to CRM, before calling the call back method
        }

        #endregion Base tool implementation

        private void btnCancel_Click(object sender, EventArgs e)
        {
            CancelWorker(); // PluginBaseControl method that calls the Background Workers CancelAsync method.

            MessageBox.Show("Cancelled");
        }

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
                chkIgnoreOwnerAndModifiedBy.Checked = true;
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
            var recordUrl = txtRecordURL.Text;

            var entity = service.Retrieve(recordUrl, new Guid(""), new ColumnSet(true));

            if(chkIgnoreAllLookups.Checked)
            {
                var allAttributesWithOutERefs = entity.Attributes.Where(attribute => attribute.GetType().Name != "EntityReference").ToList();

                entity.Attributes = new AttributeCollection();
                entity.Attributes.AddRange(allAttributesWithOutERefs);
            }
            else
            {
                if (chkIgnoreOwnerAndModifiedBy.Checked)
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
            }

            lastTargetService.Value.Create(entity);
        }
    }
}