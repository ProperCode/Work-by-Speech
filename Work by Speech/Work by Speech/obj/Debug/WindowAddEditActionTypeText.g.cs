﻿#pragma checksum "..\..\WindowAddEditActionTypeText.xaml" "{8829d00f-11b8-4213-878b-770e8597ac16}" "A50D5ABFA756333DFD5D95ABF9352EF319953E06F359DDFDCDD1DB77124865BF"
//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

using Speech;
using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging;
using System.Windows.Media.Media3D;
using System.Windows.Media.TextFormatting;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Shell;


namespace Speech {
    
    
    /// <summary>
    /// WindowAddEditActionTypeText
    /// </summary>
    public partial class WindowAddEditActionTypeText : System.Windows.Window, System.Windows.Markup.IComponentConnector {
        
        
        #line 13 "..\..\WindowAddEditActionTypeText.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.TextBox TBtext;
        
        #line default
        #line hidden
        
        
        #line 16 "..\..\WindowAddEditActionTypeText.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Button Bok;
        
        #line default
        #line hidden
        
        
        #line 25 "..\..\WindowAddEditActionTypeText.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Button Bcancel;
        
        #line default
        #line hidden
        
        private bool _contentLoaded;
        
        /// <summary>
        /// InitializeComponent
        /// </summary>
        [System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [System.CodeDom.Compiler.GeneratedCodeAttribute("PresentationBuildTasks", "4.0.0.0")]
        public void InitializeComponent() {
            if (_contentLoaded) {
                return;
            }
            _contentLoaded = true;
            System.Uri resourceLocater = new System.Uri("/Work by Speech;component/windowaddeditactiontypetext.xaml", System.UriKind.Relative);
            
            #line 1 "..\..\WindowAddEditActionTypeText.xaml"
            System.Windows.Application.LoadComponent(this, resourceLocater);
            
            #line default
            #line hidden
        }
        
        [System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [System.CodeDom.Compiler.GeneratedCodeAttribute("PresentationBuildTasks", "4.0.0.0")]
        [System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Design", "CA1033:InterfaceMethodsShouldBeCallableByChildTypes")]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity")]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1800:DoNotCastUnnecessarily")]
        void System.Windows.Markup.IComponentConnector.Connect(int connectionId, object target) {
            switch (connectionId)
            {
            case 1:
            
            #line 9 "..\..\WindowAddEditActionTypeText.xaml"
            ((Speech.WindowAddEditActionTypeText)(target)).Loaded += new System.Windows.RoutedEventHandler(this.Window_Loaded);
            
            #line default
            #line hidden
            
            #line 10 "..\..\WindowAddEditActionTypeText.xaml"
            ((Speech.WindowAddEditActionTypeText)(target)).PreviewKeyUp += new System.Windows.Input.KeyEventHandler(this.Window_PreviewKeyUp);
            
            #line default
            #line hidden
            return;
            case 2:
            this.TBtext = ((System.Windows.Controls.TextBox)(target));
            return;
            case 3:
            this.Bok = ((System.Windows.Controls.Button)(target));
            
            #line 17 "..\..\WindowAddEditActionTypeText.xaml"
            this.Bok.Click += new System.Windows.RoutedEventHandler(this.Bok_Click);
            
            #line default
            #line hidden
            return;
            case 4:
            this.Bcancel = ((System.Windows.Controls.Button)(target));
            
            #line 26 "..\..\WindowAddEditActionTypeText.xaml"
            this.Bcancel.Click += new System.Windows.RoutedEventHandler(this.Bcancel_Click);
            
            #line default
            #line hidden
            return;
            }
            this._contentLoaded = true;
        }
    }
}

