﻿#pragma checksum "..\..\WindowRecordActions.xaml" "{8829d00f-11b8-4213-878b-770e8597ac16}" "66050720E5934FAFA0EAD59339EB45C1127DD5C8DCA7DFAC1A0B15409FA90BBD"
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
    /// WindowRecordActions
    /// </summary>
    public partial class WindowRecordActions : System.Windows.Window, System.Windows.Markup.IComponentConnector {
        
        
        #line 14 "..\..\WindowRecordActions.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.ListView LVactions;
        
        #line default
        #line hidden
        
        
        #line 38 "..\..\WindowRecordActions.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.CheckBox CHBrecord_mouse;
        
        #line default
        #line hidden
        
        
        #line 43 "..\..\WindowRecordActions.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Button Bstart;
        
        #line default
        #line hidden
        
        
        #line 54 "..\..\WindowRecordActions.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Image Iquestion;
        
        #line default
        #line hidden
        
        
        #line 57 "..\..\WindowRecordActions.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Button Badd;
        
        #line default
        #line hidden
        
        
        #line 67 "..\..\WindowRecordActions.xaml"
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
            System.Uri resourceLocater = new System.Uri("/Work by Speech;component/windowrecordactions.xaml", System.UriKind.Relative);
            
            #line 1 "..\..\WindowRecordActions.xaml"
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
            
            #line 9 "..\..\WindowRecordActions.xaml"
            ((Speech.WindowRecordActions)(target)).Loaded += new System.Windows.RoutedEventHandler(this.Window_Loaded);
            
            #line default
            #line hidden
            return;
            case 2:
            this.LVactions = ((System.Windows.Controls.ListView)(target));
            return;
            case 3:
            this.CHBrecord_mouse = ((System.Windows.Controls.CheckBox)(target));
            
            #line 39 "..\..\WindowRecordActions.xaml"
            this.CHBrecord_mouse.Checked += new System.Windows.RoutedEventHandler(this.CHBrecord_mouse_Checked);
            
            #line default
            #line hidden
            
            #line 40 "..\..\WindowRecordActions.xaml"
            this.CHBrecord_mouse.Unchecked += new System.Windows.RoutedEventHandler(this.CHBrecord_mouse_Unchecked);
            
            #line default
            #line hidden
            return;
            case 4:
            this.Bstart = ((System.Windows.Controls.Button)(target));
            
            #line 45 "..\..\WindowRecordActions.xaml"
            this.Bstart.Click += new System.Windows.RoutedEventHandler(this.Bstart_Click);
            
            #line default
            #line hidden
            return;
            case 5:
            this.Iquestion = ((System.Windows.Controls.Image)(target));
            
            #line 54 "..\..\WindowRecordActions.xaml"
            this.Iquestion.PreviewMouseUp += new System.Windows.Input.MouseButtonEventHandler(this.Iquestion_PreviewMouseUp);
            
            #line default
            #line hidden
            return;
            case 6:
            this.Badd = ((System.Windows.Controls.Button)(target));
            
            #line 59 "..\..\WindowRecordActions.xaml"
            this.Badd.Click += new System.Windows.RoutedEventHandler(this.Badd_Click);
            
            #line default
            #line hidden
            return;
            case 7:
            this.Bcancel = ((System.Windows.Controls.Button)(target));
            
            #line 69 "..\..\WindowRecordActions.xaml"
            this.Bcancel.Click += new System.Windows.RoutedEventHandler(this.Bcancel_Click);
            
            #line default
            #line hidden
            return;
            }
            this._contentLoaded = true;
        }
    }
}

