   M S B A M L   `   `   `  ����I  EWork by Speech, Version=2.2.0.0, Culture=neutral, PublicKeyToken=nullclr-namespace:SpeechSpeech  5   	        Speech.WindowChooseProgram   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   &localclr-namespace:Speech     5   	   -   5          � Title$  Choose Associated Program��5   	   $	��680��6+   $	��740��68   $��14(�6D    � WindowStartupLocation$ CenterScreen=�5	   	    � 
ResizeMode$ NoResize=�6.   .��5
      � .+�5   
   �� $��Open Programs��6   $	��707��6*   $	��570��66   $��	10,10,0,0q�6C   $
��Left=�5      $	��Top=�6.   .��5      �� .+�5      �� -   5        $��
LVprograms��5   B   
����  �$��10,10,10,10q�6-   $��14(�6T   $	��486��6b   $	��Top=�5      $	��675��6Z   $��Visible=�5      $��Disabled=�5       � 	Resources 5       �� 
TargetType �� Property �� Value��%T   '��       �� 5      " ��6    .�5      ��  � ��$
 Left=�6K   5      5       �� View 5      �� 5       �� AllowsColumnReorder . .P�5      �� $��Process Name��6-   	 �� Width$		 200��6C   
 �� CellTemplate
 5   "   �� 5   &   .R�5   *   ��  �� Orientation$ 
Horizontal=�65   .+�5   .   �� $��
0, 2, 5, 2q�5   4   $��24��6H   $��24��6S   ��5   4   �� *process_image+5   a   �� 5   .   $��Center=�5    8   ��5   8   �� *process_name+5    U   5!   +   5"   '   5$      �� 5%      $��Window Title��6-   $		 465��6C    �� DisplayMemberBinding 5&   -   �� *window_title+5'      5)      �� 5*      -   5         ��Name$ Brefresh��5*   @   
���   �
���   �$��Refresh��6R   $
��Left=�6d   $��
10,504,0,0q�6   $	��Top=�5+      $	��100��6)   $��30��65   ��5,      �� 5-       �� EndPoint$ 0.5,1&�6.    � MappingMode$ RelativeToBoundingBox=�6?    �� 
StartPoint$ 0.5,0&�6c   .��5.      �  � Color$ White��6+   � 5/      $ 	#FFCDCDCD��6+    � Offset$ 1Q�6=   50      52      53      54      �� 55   
   -   5        $	 Bok��55   8   
���   �
���   �$��OK��6E   $
��Left=�6R   $��
10,599,0,0q�6m   $	��Top=�6�   $	��100��56      $��30��6   ��57      �� 58      $ 0.5,1&�6&   $ RelativeToBoundingBox=�67   $ 0.5,0&�6[   .��59      � $ White��6#   � 5:      $ 	#FFCDCDCD��6#   $ 1Q�65   5;      5=      �� 5>   
   -   5        $ Bcancel��5>   8   
���   �
���   �$��Cancel��6I   $
��Left=�6Z   $��615,599,0,0q�6u   $	��Top=�5?      $	��100��6)   $��30��65   ��5@      �� 5A      $ 0.5,1&�6&   $ RelativeToBoundingBox=�67   $ 0.5,0&�6[   .��5B      � $ White��6#   � 5C      $ 	#FFCDCDCD��6#   $ 1Q�65   5D      5F      5G      5H      