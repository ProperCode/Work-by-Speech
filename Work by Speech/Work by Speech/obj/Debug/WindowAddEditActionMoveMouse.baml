   M S B A M L   `   `   `  ����I  EWork by Speech, Version=2.2.0.0, Culture=neutral, PublicKeyToken=nullclr-namespace:SpeechSpeech  5   	   )    #Speech.WindowAddEditActionMoveMouse   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   &localclr-namespace:Speech     5   	   -   5          � Title$  Add Action: Move Mouse��5   	   $	��290��6(   $	��554��65   $��14(�6A    � WindowStartupLocation$ CenterScreen=�6O    � 
ResizeMode$ NoResize=�5	   	   .��5
      � .+�5   
   �� $��Cursor Position��6   $	��254��6,   $	��180��68   $��	10,10,0,0q�6E   $
��Left=�5      $	��Top=�6.   .��5      �� .+�5      	� -   5         	�Name$ 
RBabsolute��5      $��	10,10,0,0q�62   .��6E   
Absolute6O   	� 5      -   5        $ 
RBrelative��5      $��	10,40,0,0q�62   .��6E   
Relative6O   �� 5      $��X:��6   $��6,70,0,0q�6%   $
��Left=�67   $	��Top=�6R   �� 5      -   5         ��Name$	 TBx��5      $��23��6'   $��	27,74,0,0q�63   $
��Left=�5      $	��Top=�6-   $	��154��6E   �� 5      -   5         ��Name$ 	Iquestion��5   P   $��images/help.png��6   $��48��61   $��
187,81,0,0q�6<   5      �� 5      $��Y:��6   $��	6,108,0,0q�6%   $
��Left=�68   $	��Top=�6S   �� 5      -   5        $	 TBy��5      $��23��6'   $��
27,114,0,0q�63   $
��Left=�5      $	��Top=�6-   $	��154��6E   5      5      �� 5   
   $��Additionally Pressed Keys��6   $	��254��66   $	��180��6B   $��
274,10,0,0q�6O   $
��Left=�5      $	��Top=�6.   .��5      �� .+�5      �� -   5         ��Name$ CHBalt��5      $
��Left=�6+   $	��Top=�6F   $��	10,10,0,0q�6^   .��6q   Alt5      �� 5       -   5        $ 
CHBcontrol��5       $
��Left=�6/   $	��Top=�6J   $��	10,40,0,0q�6b   .��6u   	Control5"      �� 5#      -	   5        $ CHBshift��5#      $
��Left=�6-   $	��Top=�6H   $��	10,70,0,0q�6`   .��6s   Shift5%      �� 5&      -
   5        $ 
CHBwindows��5&      $
��Left=�6/   $	��Top=�6J   $��
10,100,0,0q�6b   .��6v   	Windows5(      5)      5*      �� 5+   
   -   5         ��Name$	 Bok��5+   8   
���   �
���   �$��OK��6E   $
��Left=�6R   $��
11,208,0,0q�6m   $	��Top=�6�   $	��100��5,      $��30��6   ��5-      �� 5.       �� EndPoint$ 0.5,1&�6&   	 � MappingMode$	 RelativeToBoundingBox=�67   
 �� 
StartPoint$
 0.5,0&�6[   .��5/      �  � Color$ White��6#   � 50      $ 	#FFCDCDCD��6#    � Offset$ 1Q�65   51      53      �� 54   
   -   5        $ Bcancel��54   8   
���   �
���   �$��Cancel��6I   $
��Left=�6Z   $��427,208,0,0q�6u   $	��Top=�55      $	��100��6)   $��30��65   ��56      �� 57      $ 0.5,1&�6&   $	 RelativeToBoundingBox=�67   $
 0.5,0&�6[   .��58      � $ White��6#   � 59      $ 	#FFCDCDCD��6#   $ 1Q�65   5:      5<      5=      5>      