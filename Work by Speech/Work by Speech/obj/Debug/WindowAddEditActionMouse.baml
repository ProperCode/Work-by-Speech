   M S B A M L   `   `   `  ����I  EWork by Speech, Version=2.2.0.0, Culture=neutral, PublicKeyToken=nullclr-namespace:SpeechSpeech  5   	   %    Speech.WindowAddEditActionMouse   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   &localclr-namespace:Speech     5   	   -   5          � Title$  Add Mouse Action��5   	   $	��400��6"   $	��554��6/   $��14(�6;    � WindowStartupLocation$ CenterScreen=�6I    � 
ResizeMode$ NoResize=�5	   	   .��5
      � .+�5   
   �� $��Mouse Button��6   $	��254��6)   $��93��65   $��	10,10,0,0q�6A   $
��Left=�5      $	��Top=�6.   .��5      �� .+�5      	� -   5         	�Name$ RBleft��5      $��	10,10,0,0q�6.    d� 	IsChecked$
 TrueY�6A   .��6R   Left6X   	� 5      -   5        $ RBright��5      $��	10,40,0,0q�6/   .��6B   Right6I   5      5      �� 5   
   $��Action��6   $	��254��6#   $	��195��6/   $��
10,105,0,0q�6<   $
��Left=�5      $	��Top=�6.   .��5      �� .+�5      	� -   5        $ RBclick��5      $��	10,10,0,0q�6/   .��6\   Click6c   	� 5      -   5        $ RBtoggle��5      $��	10,40,0,0q�60   .��6^   Toggle6f   	� 5      -   5        $ RBhold��5      $��	10,70,0,0q�6.   .��6Z   Hold6`   	� 5      -   5        $ 	RBrelease��5      $��
10,100,0,0q�61   .��6a   	Release6j   �� 5      -   5         ��Name$ Lclick_time��5      $��Click time [ms]:��6-   $��	6,130,0,0q�6H   $
��Left=�5      $	��Top=�6,   �� 5      -	   5         ��Name$ TBclick_time��5      $��23��60   $��109,134,0,0q�6<   $
��Left=�5      $	��Top=�6-   $	��120��6E   $��75��6Q   5      5      �� 5   
   -
   5         ��Name$ GBadditionally��5      $��Additionally Pressed Keys��6+   $	��254��6N   $	��152��6Z   $��
274,10,0,0q�5       $
��Left=�6'   $	��Top=�6B   .��5!      �� .+�5"      �� -   5         ��Name$ CHBalt��5"      $
��Left=�6+   $	��Top=�6F   $��	10,10,0,0q�6^   .��6q   Alt5$      �� 5%      -   5        $ 
CHBcontrol��5%      $
��Left=�6/   $	��Top=�6J   $��	10,40,0,0q�6b   .��6u   	Control5'      �� 5(      -   5        $ CHBshift��5(      $
��Left=�6-   $	��Top=�6H   $��	10,70,0,0q�6`   .��6s   Shift5*      �� 5+      -   5        $ 
CHBwindows��5+      $
��Left=�6/   $	��Top=�6J   $��
10,100,0,0q�6b   .��6v   	Windows5-      5.      5/      �� 50   
   $��Cursor Position��6   $	��254��6,   $	��137��68   $��274,163,0,0q�6E   $
��Left=�51      $	��Top=�6.   .��52      �� .+�53      �� -   5        $ CHBchange_position��53      $
��Left=�67   $	��Top=�6R   $��	10,10,0,0q�54      .��55   D   Change cursor position first57      �� 58      -   5        $ Lx��58      $��X:��6$   $��6,38,0,0q�61   $
��Left=�6C   $	��Top=�59      �� 5:      -   5        $	 TBx��5:      $��23��6'   $��	27,42,0,0q�63   $
��Left=�5;      $	��Top=�6-   $	��154��6E   �� 5<      -   5        	 ��Name$	 	Iquestion��5=      $��images/help.png��5<      $��48��61   $��
187,46,0,0q�6<   5=   +   �� 5>      -   5        $ Ly��5>      $��Y:��6$   $��6,72,0,0q�61   $
��Left=�6C   $	��Top=�5?      �� 5@      -   5        $	 TBy��5@      $��23��6'   $��	27,76,0,0q�63   $
��Left=�5A      $	��Top=�6-   $	��154��6E   5B      5C      �� 5D   
   -   5        
 ��Name$	
 Bok��5D   8   
���   �
���   �$��OK��6E   $
��Left=�6R   $��
11,318,0,0q�6m   $	��Top=�6�   $	��100��5E      $��30��6   ��5F      �� 5G       �� EndPoint$ 0.5,1&�6&    � MappingMode$ RelativeToBoundingBox=�67    �� 
StartPoint$ 0.5,0&�6[   .��5H      �  � Color$ White��6#   � 5I      $ 	#FFCDCDCD��6#    � Offset$ 1Q�65   5J      5L      �� 5M   
   -   5        $
 Bcancel��5M   8   
���   �
���   �$��Cancel��6I   $
��Left=�6Z   $��427,318,0,0q�6u   $	��Top=�5N      $	��100��6)   $��30��65   ��5O      �� 5P      $ 0.5,1&�6&   $ RelativeToBoundingBox=�67   $ 0.5,0&�6[   .��5Q      � $ White��6#   � 5R      $ 	#FFCDCDCD��6#   $ 1Q�65   5S      5U      5V      5W      