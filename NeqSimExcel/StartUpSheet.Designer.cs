﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

#pragma warning disable 414
namespace NeqSimExcel {
    
    
    /// 
    [Microsoft.VisualStudio.Tools.Applications.Runtime.StartupObjectAttribute(1)]
    [global::System.Security.Permissions.PermissionSetAttribute(global::System.Security.Permissions.SecurityAction.Demand, Name="FullTrust")]
    public sealed partial class Sheet12 : Microsoft.Office.Tools.Excel.WorksheetBase {
        
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
        private global::System.Object missing = global::System.Type.Missing;
        
        internal Microsoft.Office.Tools.Excel.Controls.CheckBox activateOperationsCheckBox;
        
        internal Microsoft.Office.Tools.Excel.Controls.CheckBox activatePVTcheckBox;
        
        internal Microsoft.Office.Tools.Excel.Controls.CheckBox exportCheckBox;
        
        internal Microsoft.Office.Tools.Excel.Controls.CheckBox RFOCheckBox;
        
        internal Microsoft.Office.Tools.Excel.Controls.CheckBox activateFACheckBox;
        
        internal Microsoft.Office.Tools.Excel.Controls.CheckBox advancedCheckBox;
        
        internal Microsoft.Office.Tools.Excel.Controls.CheckBox activateFluidOperationCheckBox;
        
        internal Microsoft.Office.Tools.Excel.Controls.CheckBox unitOperationcheckBox;
        
        internal Microsoft.Office.Tools.Excel.Controls.CheckBox processCheckBox;
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        public Sheet12(global::Microsoft.Office.Tools.Excel.Factory factory, global::System.IServiceProvider serviceProvider) : 
                base(factory, serviceProvider, "Sheet12", "Sheet12") {
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        protected override void Initialize() {
            base.Initialize();
            Globals.Sheet12 = this;
            global::System.Windows.Forms.Application.EnableVisualStyles();
            this.InitializeCachedData();
            this.InitializeControls();
            this.InitializeComponents();
            this.InitializeData();
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        protected override void FinishInitialization() {
            this.InternalStartup();
            this.OnStartup();
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        protected override void InitializeDataBindings() {
            this.BeginInitialization();
            this.BindToData();
            this.EndInitialization();
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void InitializeCachedData() {
            if ((this.DataHost == null)) {
                return;
            }
            if (this.DataHost.IsCacheInitialized) {
                this.DataHost.FillCachedData(this);
            }
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void InitializeData() {
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void BindToData() {
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        private void StartCaching(string MemberName) {
            this.DataHost.StartCaching(this, MemberName);
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        private void StopCaching(string MemberName) {
            this.DataHost.StopCaching(this, MemberName);
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        private bool IsCached(string MemberName) {
            return this.DataHost.IsCached(this, MemberName);
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void BeginInitialization() {
            this.BeginInit();
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void EndInitialization() {
            this.EndInit();
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void InitializeControls() {
            this.activateOperationsCheckBox = new Microsoft.Office.Tools.Excel.Controls.CheckBox(Globals.Factory, this.ItemProvider, this.HostContext, "17591C4241EBBE142A818AAC121AF6D4896B41", "17591C4241EBBE142A818AAC121AF6D4896B41", this, "activateOperationsCheckBox");
            this.activatePVTcheckBox = new Microsoft.Office.Tools.Excel.Controls.CheckBox(Globals.Factory, this.ItemProvider, this.HostContext, "2B53EE6D22AEC9248F22BC5F26474D9931F682", "2B53EE6D22AEC9248F22BC5F26474D9931F682", this, "activatePVTcheckBox");
            this.exportCheckBox = new Microsoft.Office.Tools.Excel.Controls.CheckBox(Globals.Factory, this.ItemProvider, this.HostContext, "39A3F493A352D5348B438D913B90F58B426BC3", "39A3F493A352D5348B438D913B90F58B426BC3", this, "exportCheckBox");
            this.RFOCheckBox = new Microsoft.Office.Tools.Excel.Controls.CheckBox(Globals.Factory, this.ItemProvider, this.HostContext, "4E5B9BBE8412A74438348D73466C7B25D11C64", "4E5B9BBE8412A74438348D73466C7B25D11C64", this, "RFOCheckBox");
            this.activateFACheckBox = new Microsoft.Office.Tools.Excel.Controls.CheckBox(Globals.Factory, this.ItemProvider, this.HostContext, "536869FBA5AFBC541C05ACE056E3A30047B075", "536869FBA5AFBC541C05ACE056E3A30047B075", this, "activateFACheckBox");
            this.advancedCheckBox = new Microsoft.Office.Tools.Excel.Controls.CheckBox(Globals.Factory, this.ItemProvider, this.HostContext, "6B08B159D6E53F648BE6AB476A6A21F0393306", "6B08B159D6E53F648BE6AB476A6A21F0393306", this, "advancedCheckBox");
            this.activateFluidOperationCheckBox = new Microsoft.Office.Tools.Excel.Controls.CheckBox(Globals.Factory, this.ItemProvider, this.HostContext, "706FBF5B478C737424A79DC37BF150F4394617", "706FBF5B478C737424A79DC37BF150F4394617", this, "activateFluidOperationCheckBox");
            this.unitOperationcheckBox = new Microsoft.Office.Tools.Excel.Controls.CheckBox(Globals.Factory, this.ItemProvider, this.HostContext, "8A3FA4ADB8327F840398B9C58ACE605A52B4D8", "8A3FA4ADB8327F840398B9C58ACE605A52B4D8", this, "unitOperationcheckBox");
            this.processCheckBox = new Microsoft.Office.Tools.Excel.Controls.CheckBox(Globals.Factory, this.ItemProvider, this.HostContext, "9557596019B87C9421C9861D9735F45E400699", "9557596019B87C9421C9861D9735F45E400699", this, "processCheckBox");
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void InitializeComponents() {
            // 
            // activateOperationsCheckBox
            // 
            this.activateOperationsCheckBox.Name = "activateOperationsCheckBox";
            this.activateOperationsCheckBox.Text = "Operations";
            // 
            // activatePVTcheckBox
            // 
            this.activatePVTcheckBox.Name = "activatePVTcheckBox";
            this.activatePVTcheckBox.Text = "PVT";
            // 
            // exportCheckBox
            // 
            this.exportCheckBox.Name = "exportCheckBox";
            this.exportCheckBox.Text = "Import/Export";
            // 
            // RFOCheckBox
            // 
            this.RFOCheckBox.Name = "RFOCheckBox";
            this.RFOCheckBox.Text = "RFO";
            // 
            // activateFACheckBox
            // 
            this.activateFACheckBox.Name = "activateFACheckBox";
            this.activateFACheckBox.Text = "Flow Assurnce";
            // 
            // advancedCheckBox
            // 
            this.advancedCheckBox.Name = "advancedCheckBox";
            this.advancedCheckBox.Text = "Advanced";
            // 
            // activateFluidOperationCheckBox
            // 
            this.activateFluidOperationCheckBox.Name = "activateFluidOperationCheckBox";
            this.activateFluidOperationCheckBox.Text = "Fluid Operations";
            // 
            // unitOperationcheckBox
            // 
            this.unitOperationcheckBox.Name = "unitOperationcheckBox";
            this.unitOperationcheckBox.Text = "Unit Operation";
            // 
            // processCheckBox
            // 
            this.processCheckBox.Name = "processCheckBox";
            this.processCheckBox.Text = "Process";
            // 
            // Sheet12
            // 
            this.activateOperationsCheckBox.BindingContext = this.BindingContext;
            this.activatePVTcheckBox.BindingContext = this.BindingContext;
            this.exportCheckBox.BindingContext = this.BindingContext;
            this.RFOCheckBox.BindingContext = this.BindingContext;
            this.activateFACheckBox.BindingContext = this.BindingContext;
            this.advancedCheckBox.BindingContext = this.BindingContext;
            this.activateFluidOperationCheckBox.BindingContext = this.BindingContext;
            this.unitOperationcheckBox.BindingContext = this.BindingContext;
            this.processCheckBox.BindingContext = this.BindingContext;
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        private bool NeedsFill(string MemberName) {
            return this.DataHost.NeedsFill(this, MemberName);
        }
    }
    
    internal sealed partial class Globals {
        
        private static Sheet12 _Sheet12;
        
        internal static Sheet12 Sheet12 {
            get {
                return _Sheet12;
            }
            set {
                if ((_Sheet12 == null)) {
                    _Sheet12 = value;
                }
                else {
                    throw new System.NotSupportedException();
                }
            }
        }
    }
}
