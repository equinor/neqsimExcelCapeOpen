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
    [Microsoft.VisualStudio.Tools.Applications.Runtime.StartupObjectAttribute(2)]
    [global::System.Security.Permissions.PermissionSetAttribute(global::System.Security.Permissions.SecurityAction.Demand, Name="FullTrust")]
    public sealed partial class Sheet1 : Microsoft.Office.Tools.Excel.WorksheetBase {
        
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "14.0.0.0")]
        private global::System.Object missing = global::System.Type.Missing;
        
        internal Microsoft.Office.Tools.Excel.Controls.Button button1;
        
        internal Microsoft.Office.Tools.Excel.Controls.Button normalizeButton;
        
        internal Microsoft.Office.Tools.Excel.Controls.Button clearButton;
        
        internal Microsoft.Office.Tools.Excel.Controls.Button addComponentButoon;
        
        internal Microsoft.Office.Tools.Excel.Controls.Button molToWtButton;
        
        internal Microsoft.Office.Tools.Excel.Controls.RadioButton molRadioButton;
        
        internal Microsoft.Office.Tools.Excel.Controls.RadioButton wtRadioButton;
        
        internal Microsoft.Office.Tools.Excel.Controls.RadioButton plusFracRadioButton;
        
        internal Microsoft.Office.Tools.Excel.Controls.RadioButton noPlusRadioButton;
        
        internal Microsoft.Office.Tools.Excel.Controls.CheckBox MWcheckBox;
        
        internal Microsoft.Office.Tools.Excel.Controls.CheckBox watsatcheckBox;
        
        internal Microsoft.Office.Tools.Excel.Controls.Button tuneOkbutton;
        
        internal Microsoft.Office.Tools.Excel.Controls.ComboBox EoScombobox;
        
        internal Microsoft.Office.Tools.Excel.Controls.Button button2;
        
        internal Microsoft.Office.Tools.Excel.Controls.ComboBox inhibitorComboBox;
        
        internal Microsoft.Office.Tools.Excel.Controls.CheckBox waxcheckBox;
        
        internal Microsoft.Office.Tools.Excel.Controls.CheckBox solidGlycolCheckBox;
        
        internal Microsoft.Office.Tools.Excel.Controls.CheckBox iceCheckBox;
        
        internal Microsoft.Office.Tools.Excel.Controls.CheckBox solidCO2checkBox;
        
        internal Microsoft.Office.Tools.Excel.Controls.ComboBox numberOfPseudoCompComboBox;
        
        internal Microsoft.Office.Tools.Excel.Controls.Label numbPseudoLabel;
        
        internal Microsoft.Office.Tools.Excel.Controls.ComboBox inhibitorCalcTypecomboBox;
        
        internal Microsoft.Office.Tools.Excel.Controls.CheckBox chemicalReactionsCheckBox;
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        public Sheet1(global::Microsoft.Office.Tools.Excel.Factory factory, global::System.IServiceProvider serviceProvider) : 
                base(factory, serviceProvider, "Sheet1", "Sheet1") {
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "14.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        protected override void Initialize() {
            base.Initialize();
            Globals.Sheet1 = this;
            global::System.Windows.Forms.Application.EnableVisualStyles();
            this.InitializeCachedData();
            this.InitializeControls();
            this.InitializeComponents();
            this.InitializeData();
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "14.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        protected override void FinishInitialization() {
            this.InternalStartup();
            this.OnStartup();
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "14.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        protected override void InitializeDataBindings() {
            this.BeginInitialization();
            this.BindToData();
            this.EndInitialization();
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "14.0.0.0")]
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
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "14.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void InitializeData() {
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "14.0.0.0")]
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
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "14.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void BeginInitialization() {
            this.BeginInit();
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "14.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void EndInitialization() {
            this.EndInit();
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "14.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void InitializeControls() {
            this.button1 = new Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, this.ItemProvider, this.HostContext, "2DB16E32A2195924F98291032DA20708D53D52", "2DB16E32A2195924F98291032DA20708D53D52", this, "button1");
            this.normalizeButton = new Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, this.ItemProvider, this.HostContext, "458FBF677492C1444B2490E74341E048866584", "458FBF677492C1444B2490E74341E048866584", this, "normalizeButton");
            this.clearButton = new Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, this.ItemProvider, this.HostContext, "59D3618C85F26C547E558CA85FE14EE2FCF4E5", "59D3618C85F26C547E558CA85FE14EE2FCF4E5", this, "clearButton");
            this.addComponentButoon = new Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, this.ItemProvider, this.HostContext, "6CA29245266C5564D8069D6065AE9EF09F4DA6", "6CA29245266C5564D8069D6065AE9EF09F4DA6", this, "addComponentButoon");
            this.molToWtButton = new Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, this.ItemProvider, this.HostContext, "73EBFB3597D7A37487A7B6FB76FEE3F98877F7", "73EBFB3597D7A37487A7B6FB76FEE3F98877F7", this, "molToWtButton");
            this.molRadioButton = new Microsoft.Office.Tools.Excel.Controls.RadioButton(Globals.Factory, this.ItemProvider, this.HostContext, "8578E6734835BB84C2D886DE8DC04D61281918", "8578E6734835BB84C2D886DE8DC04D61281918", this, "molRadioButton");
            this.wtRadioButton = new Microsoft.Office.Tools.Excel.Controls.RadioButton(Globals.Factory, this.ItemProvider, this.HostContext, "99E80A20B9DC9A94CD39B2759AD8F94B050069", "99E80A20B9DC9A94CD39B2759AD8F94B050069", this, "wtRadioButton");
            this.plusFracRadioButton = new Microsoft.Office.Tools.Excel.Controls.RadioButton(Globals.Factory, this.ItemProvider, this.HostContext, "19357786F1D686149141A459153C0DF2C0F0C1", "19357786F1D686149141A459153C0DF2C0F0C1", this, "plusFracRadioButton");
            this.noPlusRadioButton = new Microsoft.Office.Tools.Excel.Controls.RadioButton(Globals.Factory, this.ItemProvider, this.HostContext, "1029F6A33115A9145641BB711EAF71403AF5C1", "1029F6A33115A9145641BB711EAF71403AF5C1", this, "noPlusRadioButton");
            this.MWcheckBox = new Microsoft.Office.Tools.Excel.Controls.CheckBox(Globals.Factory, this.ItemProvider, this.HostContext, "184F374311CE991474D19A171449FEEFEF1BE1", "184F374311CE991474D19A171449FEEFEF1BE1", this, "MWcheckBox");
            this.watsatcheckBox = new Microsoft.Office.Tools.Excel.Controls.CheckBox(Globals.Factory, this.ItemProvider, this.HostContext, "14D36455D191EC140E31BD561E321CC6D7B5F1", "14D36455D191EC140E31BD561E321CC6D7B5F1", this, "watsatcheckBox");
            this.tuneOkbutton = new Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, this.ItemProvider, this.HostContext, "1994EA02812BDD14C0B1A8FB1F4B95D2BB2B01", "1994EA02812BDD14C0B1A8FB1F4B95D2BB2B01", this, "tuneOkbutton");
            this.EoScombobox = new Microsoft.Office.Tools.Excel.Controls.ComboBox(Globals.Factory, this.ItemProvider, this.HostContext, "1D2E9956B12741147DD1A3DF1201F2754A13C1", "1D2E9956B12741147DD1A3DF1201F2754A13C1", this, "EoScombobox");
            this.button2 = new Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, this.ItemProvider, this.HostContext, "11D1A8EB214D8514C5B1A3D31F8935F5091591", "11D1A8EB214D8514C5B1A3D31F8935F5091591", this, "button2");
            this.inhibitorComboBox = new Microsoft.Office.Tools.Excel.Controls.ComboBox(Globals.Factory, this.ItemProvider, this.HostContext, "11F29B8CD1F51914BFA19C4D1DC2D2C131B9A1", "11F29B8CD1F51914BFA19C4D1DC2D2C131B9A1", this, "inhibitorComboBox");
            this.waxcheckBox = new Microsoft.Office.Tools.Excel.Controls.CheckBox(Globals.Factory, this.ItemProvider, this.HostContext, "1EAAA176C1AC7514C061868B15C05B5B35E0B1", "1EAAA176C1AC7514C061868B15C05B5B35E0B1", this, "waxcheckBox");
            this.solidGlycolCheckBox = new Microsoft.Office.Tools.Excel.Controls.CheckBox(Globals.Factory, this.ItemProvider, this.HostContext, "23A4DCD2324F1824D612809F27465A32D2B212", "23A4DCD2324F1824D612809F27465A32D2B212", this, "solidGlycolCheckBox");
            this.iceCheckBox = new Microsoft.Office.Tools.Excel.Controls.CheckBox(Globals.Factory, this.ItemProvider, this.HostContext, "24048C03D29F45243022A2762251F47955D4B2", "24048C03D29F45243022A2762251F47955D4B2", this, "iceCheckBox");
            this.solidCO2checkBox = new Microsoft.Office.Tools.Excel.Controls.CheckBox(Globals.Factory, this.ItemProvider, this.HostContext, "226259A8A2EDB22432E2A3D52F494200BC9342", "226259A8A2EDB22432E2A3D52F494200BC9342", this, "solidCO2checkBox");
            this.numberOfPseudoCompComboBox = new Microsoft.Office.Tools.Excel.Controls.ComboBox(Globals.Factory, this.ItemProvider, this.HostContext, "27563C9BF25B55242AF2B7762A82920C774BC2", "27563C9BF25B55242AF2B7762A82920C774BC2", this, "numberOfPseudoCompComboBox");
            this.numbPseudoLabel = new Microsoft.Office.Tools.Excel.Controls.Label(Globals.Factory, this.ItemProvider, this.HostContext, "277E7FC0E21783246212B58B2585CA2DA9FA62", "277E7FC0E21783246212B58B2585CA2DA9FA62", this, "numbPseudoLabel");
            this.inhibitorCalcTypecomboBox = new Microsoft.Office.Tools.Excel.Controls.ComboBox(Globals.Factory, this.ItemProvider, this.HostContext, "291D019922DD0C241822A1892E3343DA966B02", "291D019922DD0C241822A1892E3343DA966B02", this, "inhibitorCalcTypecomboBox");
            this.chemicalReactionsCheckBox = new Microsoft.Office.Tools.Excel.Controls.CheckBox(Globals.Factory, this.ItemProvider, this.HostContext, "204D97F042C4FD24B3E2904A251AE0D274D8F2", "204D97F042C4FD24B3E2904A251AE0D274D8F2", this, "chemicalReactionsCheckBox");
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "14.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void InitializeComponents() {
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.SystemColors.Control;
            this.button1.ForeColor = System.Drawing.SystemColors.ControlText;
            this.button1.Name = "button1";
            this.button1.Text = "Ok";
            this.button1.UseVisualStyleBackColor = false;
            // 
            // normalizeButton
            // 
            this.normalizeButton.BackColor = System.Drawing.SystemColors.Control;
            this.normalizeButton.ForeColor = System.Drawing.SystemColors.ControlText;
            this.normalizeButton.Name = "normalizeButton";
            this.normalizeButton.Text = "Normalize";
            this.normalizeButton.UseVisualStyleBackColor = false;
            // 
            // clearButton
            // 
            this.clearButton.BackColor = System.Drawing.SystemColors.Control;
            this.clearButton.ForeColor = System.Drawing.SystemColors.ControlText;
            this.clearButton.Name = "clearButton";
            this.clearButton.Text = "Clear";
            this.clearButton.UseVisualStyleBackColor = false;
            // 
            // addComponentButoon
            // 
            this.addComponentButoon.BackColor = System.Drawing.SystemColors.Control;
            this.addComponentButoon.ForeColor = System.Drawing.SystemColors.ControlText;
            this.addComponentButoon.Name = "addComponentButoon";
            this.addComponentButoon.Text = "Add Comps";
            this.addComponentButoon.UseVisualStyleBackColor = false;
            // 
            // molToWtButton
            // 
            this.molToWtButton.BackColor = System.Drawing.SystemColors.Control;
            this.molToWtButton.ForeColor = System.Drawing.SystemColors.ControlText;
            this.molToWtButton.Name = "molToWtButton";
            this.molToWtButton.Text = "Mol to Weight";
            this.molToWtButton.UseVisualStyleBackColor = false;
            // 
            // molRadioButton
            // 
            this.molRadioButton.Checked = true;
            this.molRadioButton.Name = "molRadioButton";
            this.molRadioButton.Text = "mol%";
            // 
            // wtRadioButton
            // 
            this.wtRadioButton.Name = "wtRadioButton";
            this.wtRadioButton.Text = "wt%";
            // 
            // plusFracRadioButton
            // 
            this.plusFracRadioButton.Name = "plusFracRadioButton";
            this.plusFracRadioButton.Text = "Plus fraction";
            // 
            // noPlusRadioButton
            // 
            this.noPlusRadioButton.Checked = true;
            this.noPlusRadioButton.Name = "noPlusRadioButton";
            this.noPlusRadioButton.Text = "No Plus fraction";
            // 
            // MWcheckBox
            // 
            this.MWcheckBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F);
            this.MWcheckBox.Name = "MWcheckBox";
            this.MWcheckBox.Text = "Mw plus";
            // 
            // watsatcheckBox
            // 
            this.watsatcheckBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F);
            this.watsatcheckBox.Name = "watsatcheckBox";
            this.watsatcheckBox.Text = "Water saturate";
            // 
            // tuneOkbutton
            // 
            this.tuneOkbutton.BackColor = System.Drawing.SystemColors.Control;
            this.tuneOkbutton.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.5F);
            this.tuneOkbutton.ForeColor = System.Drawing.SystemColors.ControlText;
            this.tuneOkbutton.Name = "tuneOkbutton";
            this.tuneOkbutton.Text = "Ok";
            this.tuneOkbutton.UseVisualStyleBackColor = false;
            // 
            // EoScombobox
            // 
            this.EoScombobox.Items.AddRange(new object[] {
                        "Automatic",
                        "SRK-EOS",
                        "PR-EOS",
                        "CPAs-SRK-EOS-statoil",
                        "UMR-PRU-EoS"});
            this.EoScombobox.Name = "EoScombobox";
            this.EoScombobox.Text = "Automatic";
            // 
            // button2
            // 
            this.button2.BackColor = System.Drawing.SystemColors.Control;
            this.button2.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.5F);
            this.button2.ForeColor = System.Drawing.SystemColors.ControlText;
            this.button2.Name = "button2";
            this.button2.Text = "Ok";
            this.button2.UseVisualStyleBackColor = false;
            // 
            // inhibitorComboBox
            // 
            this.inhibitorComboBox.Items.AddRange(new object[] {
                        "methanol",
                        "ethanol",
                        "MEG",
                        "TEG"});
            this.inhibitorComboBox.Name = "inhibitorComboBox";
            this.inhibitorComboBox.Text = "MEG";
            // 
            // waxcheckBox
            // 
            this.waxcheckBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F);
            this.waxcheckBox.Name = "waxcheckBox";
            this.waxcheckBox.Text = "Wax";
            // 
            // solidGlycolCheckBox
            // 
            this.solidGlycolCheckBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F);
            this.solidGlycolCheckBox.Name = "solidGlycolCheckBox";
            this.solidGlycolCheckBox.Text = "Solid glycol";
            // 
            // iceCheckBox
            // 
            this.iceCheckBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F);
            this.iceCheckBox.Name = "iceCheckBox";
            this.iceCheckBox.Text = "Ice";
            // 
            // solidCO2checkBox
            // 
            this.solidCO2checkBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F);
            this.solidCO2checkBox.Name = "solidCO2checkBox";
            this.solidCO2checkBox.Text = "CO2";
            // 
            // numberOfPseudoCompComboBox
            // 
            this.numberOfPseudoCompComboBox.Items.AddRange(new object[] {
                        "5",
                        "6",
                        "7",
                        "8",
                        "9",
                        "10",
                        "11",
                        "12",
                        "13",
                        "14",
                        "15",
                        "16",
                        "17",
                        "18",
                        "19",
                        "20"});
            this.numberOfPseudoCompComboBox.Name = "numberOfPseudoCompComboBox";
            this.numberOfPseudoCompComboBox.Text = "7";
            // 
            // numbPseudoLabel
            // 
            this.numbPseudoLabel.Name = "numbPseudoLabel";
            this.numbPseudoLabel.Text = "Pseudo numb";
            // 
            // inhibitorCalcTypecomboBox
            // 
            this.inhibitorCalcTypecomboBox.Items.AddRange(new object[] {
                        "estimate wt%",
                        "set wt%"});
            this.inhibitorCalcTypecomboBox.Name = "inhibitorCalcTypecomboBox";
            this.inhibitorCalcTypecomboBox.Text = "estimate wt%";
            // 
            // chemicalReactionsCheckBox
            // 
            this.chemicalReactionsCheckBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F);
            this.chemicalReactionsCheckBox.Name = "chemicalReactionsCheckBox";
            this.chemicalReactionsCheckBox.Text = "Add chemical reactions";
            // 
            // Sheet1
            // 
            this.button1.BindingContext = this.BindingContext;
            this.normalizeButton.BindingContext = this.BindingContext;
            this.clearButton.BindingContext = this.BindingContext;
            this.addComponentButoon.BindingContext = this.BindingContext;
            this.molToWtButton.BindingContext = this.BindingContext;
            this.molRadioButton.BindingContext = this.BindingContext;
            this.wtRadioButton.BindingContext = this.BindingContext;
            this.plusFracRadioButton.BindingContext = this.BindingContext;
            this.noPlusRadioButton.BindingContext = this.BindingContext;
            this.MWcheckBox.BindingContext = this.BindingContext;
            this.watsatcheckBox.BindingContext = this.BindingContext;
            this.tuneOkbutton.BindingContext = this.BindingContext;
            this.EoScombobox.BindingContext = this.BindingContext;
            this.button2.BindingContext = this.BindingContext;
            this.inhibitorComboBox.BindingContext = this.BindingContext;
            this.waxcheckBox.BindingContext = this.BindingContext;
            this.solidGlycolCheckBox.BindingContext = this.BindingContext;
            this.iceCheckBox.BindingContext = this.BindingContext;
            this.solidCO2checkBox.BindingContext = this.BindingContext;
            this.numberOfPseudoCompComboBox.BindingContext = this.BindingContext;
            this.numbPseudoLabel.BindingContext = this.BindingContext;
            this.inhibitorCalcTypecomboBox.BindingContext = this.BindingContext;
            this.chemicalReactionsCheckBox.BindingContext = this.BindingContext;
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        private bool NeedsFill(string MemberName) {
            return this.DataHost.NeedsFill(this, MemberName);
        }
    }
    
    internal sealed partial class Globals {
        
        private static Sheet1 _Sheet1;
        
        internal static Sheet1 Sheet1 {
            get {
                return _Sheet1;
            }
            set {
                if ((_Sheet1 == null)) {
                    _Sheet1 = value;
                }
                else {
                    throw new System.NotSupportedException();
                }
            }
        }
    }
}
