   M S B A M L   `   `   `  ����I  EWork by Speech, Version=2.2.0.0, Culture=neutral, PublicKeyToken=nullclr-namespace:SpeechSpeech  5   	        Speech.WindowRecordActions   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   &localclr-namespace:Speech     5   	   -   5          � Title$  Record Macro��5   	   $	��680��6   $	��743��6+   $��14(�67    � WindowStartupLocation$ CenterScreen=�5	   	    � 
ResizeMode$ NoResize=�6.   .��5
      � .+�5   
   �� $��Recorded Actions��6   $	��707��6-   $	��570��69   $��	10,10,0,0q�6F   $
��Left=�5      $	��Top=�6.   .��5      �� .+�5      �� -   5        $��	LVactions��5   B   
����  �$��10,10,10,10q�6-   $��14(�6S   $	��456��6a   $	��Top=�5      $	��675��63   $��Visible=�5      $��Disabled=�5       � 	Resources 5       �� 
TargetType �� Property �� Value��%T   '��       �� 5      " ��6    .�5      ��  � ��$
 Left=�6K   5      5       �� View 5      �� 5       �� AllowsColumnReorder . .P�5      �� $��Action��6-   	 �� Width$		 665��6=   
 �� CellTemplate
 5   "   �� 5   &   .R�5   *   ��  �� Orientation$ 
Horizontal=�65   .+�5   .   �� $��Center=�5   8   ��5   8   �� *action+5   U   5   +   5    '   5"      5#      5%      �� 5&      -   5         ��Name$ CHBrecord_mouse��5&      $
��Left=�64   $	��Top=�6O   $��
10,475,0,0q�5'      .��5(   A   Record mouse movements5*      �� 5+      -   5         ��Name$ Bstart��5+   @   
���   �
���   �$��Start��6P   $
��Left=�5,      $��
10,504,0,0q�64   $	��Top=�5-      $	��100��61   $��30��6=   ��5.      �� 5/       �� EndPoint$ 0.5,1&�6.    � MappingMode$ RelativeToBoundingBox=�6?    �� 
StartPoint$ 0.5,0&�6c   .��50      �  � Color$ White��6+   � 51      $ 	#FFCDCDCD��6+    � Offset$ 1Q�6=   52      54      �� 55      -   5         ��Name$ 	Iquestion��56   -   $��images/help.png��55      $��30��61   $��115,504,0,0q�56      6@   57      58      �� 59   
   -   5        $
 Badd��59   8   
���   �
���   �$	��Add��6F   $
��Left=�5:      $��
11,596,0,0q�6,   $	��Top=�6@   $	��100��5;      $��30��6   {�.  ��5<      �� 5=      $ 0.5,1&�6&   $ RelativeToBoundingBox=�67   $ 0.5,0&�6[   .��5>      � $ White��6#   � 5?      $ 	#FFCDCDCD��6#   $ 1Q�65   5@      5B      �� 5C   
   -   5        $ Bcancel��5C   8   
���   �
���   �$��Cancel��6I   $
��Left=�5D      $��616,596,0,0q�6,   $	��Top=�5E      $	��100��6)   $��30��65   ��5F      �� 5G      $ 0.5,1&�6&   $ RelativeToBoundingBox=�67   $ 0.5,0&�6[   .��5H      � $ White��6#   � 5I      $ 	#FFCDCDCD��6#   $ 1Q�65   5J      5L      5M      5N      