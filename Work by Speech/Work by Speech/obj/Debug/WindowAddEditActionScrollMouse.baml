   M S B A M L   `   `   `  ����I  EWork by Speech, Version=2.2.0.0, Culture=neutral, PublicKeyToken=nullclr-namespace:SpeechSpeech  5   	   +    %Speech.WindowAddEditActionScrollMouse   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   &localclr-namespace:Speech     5   	   -   5          � Title$$  Add Action: Scroll Mouse Wheel��5   	   $	��262��60   $	��554��6=   $��14(�6I    � WindowStartupLocation$ CenterScreen=�5	   	    � 
ResizeMode$ NoResize=�6.   .��5      � .+�5   
   �� $��Scrolling Settings��6   $	��254��6/   $	��152��6;   $��	10,10,0,0q�6H   $
��Left=�5      $	��Top=�6.   .��5      �� .+�5      	� -   5         	�Name$ RBscroll_up��5      $��	10,10,0,0q�63   .��6F   	Scroll up6Q   	� 5      -   5        $ RBscroll_down��5      $��	10,40,0,0q�65   .��6H   Scroll down6U   �� 5      $��Scrolling value:��6   $��6,70,0,0q�63   $
��Left=�6E   $	��Top=�6`   �� 5      -   5         ��Name$ TBscrolling_value��5      $��23��65   $��
108,74,0,0q�6A   $
��Left=�5      $	��Top=�6-   $	��119��6E   5      5      �� 5   
   $��Additionally Pressed Keys��6   $	��254��66   $	��152��6B   $��
274,10,0,0q�6O   $
��Left=�5      $	��Top=�6.   .��5      �� .+�5      �� -   5         ��Name$ CHBalt��5      $
��Left=�6+   $	��Top=�6F   $��	10,10,0,0q�6^   .��6q   Alt5      �� 5      -   5        $ 
CHBcontrol��5      $
��Left=�6/   $	��Top=�6J   $��	10,40,0,0q�6b   .��6u   	Control5      �� 5      -   5        $ CHBshift��5      $
��Left=�6-   $	��Top=�6H   $��	10,70,0,0q�6`   .��6s   Shift5!      �� 5"      -   5        $ 
CHBwindows��5"      $
��Left=�6/   $	��Top=�6J   $��
10,100,0,0q�6b   .��6v   	Windows5$      5%      5&      �� 5'   
   -	   5         ��Name$	 Bok��5'   8   
���   �
���   �$��OK��6E   $
��Left=�6R   $��
11,180,0,0q�6m   $	��Top=�6�   $	��100��5(      $��30��6   ��5)      �� 5*       �� EndPoint$ 0.5,1&�6&    � MappingMode$ RelativeToBoundingBox=�67   	 �� 
StartPoint$	 0.5,0&�6[   .��5+      � 
 � Color$
 White��6#   � 5,      $
 	#FFCDCDCD��6#    � Offset$ 1Q�65   5-      5/      �� 50   
   -
   5        $ Bcancel��50   8   
���   �
���   �$��Cancel��6I   $
��Left=�6Z   $��427,180,0,0q�6u   $	��Top=�51      $	��100��6)   $��30��65   ��52      �� 53      $ 0.5,1&�6&   $ RelativeToBoundingBox=�67   $	 0.5,0&�6[   .��54      � $
 White��6#   � 55      $
 	#FFCDCDCD��6#   $ 1Q�65   56      58      59      5:      