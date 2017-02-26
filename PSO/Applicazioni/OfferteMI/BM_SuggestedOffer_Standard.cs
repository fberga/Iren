﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

using System.Xml.Serialization;
namespace Iren.PSO.Applicazioni
{ 
// 
// This source code was auto-generated by xsd, Version=4.6.1055.0.
// 
/// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.6.1055.0")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true, Namespace="urn:XML-BIDMGM")]
    [System.Xml.Serialization.XmlRootAttribute(Namespace="urn:XML-BIDMGM", IsNullable=false)]

    public partial class Suggested {   
        private SuggestedCoordinate[] coordinateField; 
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("Coordinate")]
        public SuggestedCoordinate[] Coordinate {
            get {
                return this.coordinateField;
            }
            set {
                this.coordinateField = value;
            }
        }
    }
/// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.6.1055.0")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true, Namespace="urn:XML-BIDMGM")]
    public partial class SuggestedCoordinate {  
        private SuggestedCoordinateSG1[] sG1Field; 
        private SuggestedCoordinateSG2[] sG2Field; 
        private SuggestedCoordinateSG3[] sG3Field;
        private SuggestedCoordinateSG4[] sG4Field;
        private ElencoMercatiEnergia mercatoField;
        private string iDUnitField;
        private string flowDateField;
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("SG1")]
        public SuggestedCoordinateSG1[] SG1 {
            get {
                return this.sG1Field;
            }
            set {
                this.sG1Field = value;
            }
        }
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("SG2")]
        public SuggestedCoordinateSG2[] SG2 {
            get {
                return this.sG2Field;
            }
            set {
                this.sG2Field = value;
            }
        }
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("SG3")]
        public SuggestedCoordinateSG3[] SG3 {
            get {
                return this.sG3Field;
            }
            set {
                this.sG3Field = value;
            }
        }  
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("SG4")]
        public SuggestedCoordinateSG4[] SG4 {
            get {
                return this.sG4Field;
            }
            set {
                this.sG4Field = value;
            }
        }  
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public ElencoMercatiEnergia Mercato {
            get {
                return this.mercatoField;
            }
            set {
                this.mercatoField = value;
            }
        }
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string IDUnit {
            get {
                return this.iDUnitField;
            }
            set {
                this.iDUnitField = value;
            }
        }
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(DataType="integer")]
        public string FlowDate {
            get {
                return this.flowDateField;
            }
            set {
                this.flowDateField = value;
            }
        }
    }
/// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.6.1055.0")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true, Namespace="urn:XML-BIDMGM")]
    public partial class SuggestedCoordinateSG1 { 
        private string qUAField;
        private string pREField;
        private string bILANCField;
        private TipoAzione aZIONEField;
        private bool aZIONEFieldSpecified;
        private string valueField;
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string QUA {
            get {
                return this.qUAField;
            }
            set {
                this.qUAField = value;
            }
        }
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string PRE {
            get {
                return this.pREField;
            }
            set {
                this.pREField = value;
            }
        }
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string BILANC {
            get {
                return this.bILANCField;
            }
            set {
                this.bILANCField = value;
            }
        }
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public TipoAzione AZIONE {
            get {
                return this.aZIONEField;
            }
            set {
                this.aZIONEField = value;
            }
        }
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool AZIONESpecified {
            get {
                return this.aZIONEFieldSpecified;
            }
            set {
                this.aZIONEFieldSpecified = value;
            }
        }
        /// <remarks/>
        [System.Xml.Serialization.XmlTextAttribute(DataType="integer")]
        public string Value {
            get {
                return this.valueField;
            }
            set {
                this.valueField = value;
            }
        }
    }
/// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.6.1055.0")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="urn:XML-BIDMGM")]
    public enum TipoAzione { 
        /// <remarks/>
        ACQ,
        /// <remarks/>
        VEN,
    }
/// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.6.1055.0")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true, Namespace="urn:XML-BIDMGM")]
    public partial class SuggestedCoordinateSG2 {
        private string qUAField;
        private string pREField;
        private string bILANCField;
        private TipoAzione aZIONEField;
        private bool aZIONEFieldSpecified;
        private string valueField;
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string QUA {
            get {
                return this.qUAField;
            }
            set {
                this.qUAField = value;
            }
        }
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string PRE {
            get {
                return this.pREField;
            }
            set {
                this.pREField = value;
            }
        }
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string BILANC {
            get {
                return this.bILANCField;
            }
            set {
                this.bILANCField = value;
            }
        }
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public TipoAzione AZIONE {
            get {
                return this.aZIONEField;
            }
            set {
                this.aZIONEField = value;
            }
        }
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool AZIONESpecified {
            get {
                return this.aZIONEFieldSpecified;
            }
            set {
                this.aZIONEFieldSpecified = value;
            }
        }
        /// <remarks/>
        [System.Xml.Serialization.XmlTextAttribute(DataType="integer")]
        public string Value {
            get {
                return this.valueField;
            }
            set {
                this.valueField = value;
            }
        }
    }
/// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.6.1055.0")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true, Namespace="urn:XML-BIDMGM")]
    public partial class SuggestedCoordinateSG3 {
        private string qUAField;
        private string pREField;
        private string bILANCField;
        private TipoAzione aZIONEField;
        private bool aZIONEFieldSpecified;
        private string valueField;
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string QUA {
            get {
                return this.qUAField;
            }
            set {
                this.qUAField = value;
            }
        }
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string PRE {
            get {
                return this.pREField;
            }
            set {
                this.pREField = value;
            }
        }
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string BILANC {
            get {
                return this.bILANCField;
            }
            set {
                this.bILANCField = value;
            }
        }
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public TipoAzione AZIONE {
            get {
                return this.aZIONEField;
            }
            set {
                this.aZIONEField = value;
            }
        }
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool AZIONESpecified {
            get {
                return this.aZIONEFieldSpecified;
            }
            set {
                this.aZIONEFieldSpecified = value;
            }
        }
        /// <remarks/>
        [System.Xml.Serialization.XmlTextAttribute(DataType="integer")]
        public string Value {
            get {
                return this.valueField;
            }
            set {
                this.valueField = value;
            }
        }
    }
/// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.6.1055.0")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true, Namespace="urn:XML-BIDMGM")]
    public partial class SuggestedCoordinateSG4 {
        private string qUAField;
        private string pREField;
        private string bILANCField;
        private TipoAzione aZIONEField;
        private bool aZIONEFieldSpecified;
        private string valueField;
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string QUA {
            get {
                return this.qUAField;
            }
            set {
                this.qUAField = value;
            }
        }
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string PRE {
            get {
                return this.pREField;
            }
            set {
                this.pREField = value;
            }
        }
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string BILANC {
            get {
                return this.bILANCField;
            }
            set {
                this.bILANCField = value;
            }
        }
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public TipoAzione AZIONE {
            get {
                return this.aZIONEField;
            }
            set {
                this.aZIONEField = value;
            }
        }
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool AZIONESpecified {
            get {
                return this.aZIONEFieldSpecified;
            }
            set {
                this.aZIONEFieldSpecified = value;
            }
        }
        /// <remarks/>
        [System.Xml.Serialization.XmlTextAttribute(DataType="integer")]
        public string Value {
            get {
                return this.valueField;
            }
            set {
                this.valueField = value;
            }
        }
    }
/// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.6.1055.0")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="urn:XML-BIDMGM")]
    public enum ElencoMercatiEnergia {
        /// <remarks/>
        PCE,
        /// <remarks/>
        MGP,
        /// <remarks/>
        MA1,
        /// <remarks/>
        MI1,
        /// <remarks/>
        MI2,
        /// <remarks/>
        MI3,
        /// <remarks/>
        MI4,
        /// <remarks/>
        MI5,
        /// <remarks/>
        MI6,
        /// <remarks/>
        MI7,
    }
/// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.6.1055.0")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true, Namespace="urn:XML-BIDMGM")]
    [System.Xml.Serialization.XmlRootAttribute("BMTransaction-SUG", Namespace="urn:XML-BIDMGM", IsNullable=false)]
    public partial class BMTransactionSUG {
        private Suggested suggestedField;
        private string referenceNumberField;
        private string dataCreazioneField;
        private string oraCreazioneField;       
        private YESNO applySendAutomaticField;
        private bool applySendAutomaticFieldSpecified;
        private string operatorCreatorField;
        /// <remarks/>
        public Suggested Suggested {
            get {
                return this.suggestedField;
            }
            set {
                this.suggestedField = value;
            }
        }
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string ReferenceNumber {
            get {
                return this.referenceNumberField;
            }
            set {
                this.referenceNumberField = value;
            }
        }
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(DataType="integer")]
        public string DataCreazione {
            get {
                return this.dataCreazioneField;
            }
            set {
                this.dataCreazioneField = value;
            }
        }
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute(DataType="integer")]
        public string OraCreazione {
            get {
                return this.oraCreazioneField;
            }
            set {
                this.oraCreazioneField = value;
            }
        }       
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public YESNO ApplySendAutomatic {
            get {
                return this.applySendAutomaticField;
            }
            set {
                this.applySendAutomaticField = value;
            }
        }     
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool ApplySendAutomaticSpecified {
            get {
                return this.applySendAutomaticFieldSpecified;
            }
            set {
                this.applySendAutomaticFieldSpecified = value;
            }
        }  
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string OperatorCreator {
            get {
                return this.operatorCreatorField;
            }
            set {
                this.operatorCreatorField = value;
            }
        }
    }
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.6.1055.0")]
    [System.SerializableAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="urn:XML-BIDMGM")]
    public enum YESNO {    
        /// <remarks/>
        YES,     
        /// <remarks/>
        NO,
    }
}
