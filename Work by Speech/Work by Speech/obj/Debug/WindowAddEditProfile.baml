   M S B A M L   `   `   `  ����I  EWork by Speech, Version=2.2.0.0, Culture=neutral, PublicKeyToken=nullclr-namespace:SpeechSpeech  5   	   !    Speech.WindowAddEditProfile   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   &localclr-namespace:Speech     5   	   -   5          � Title$  Add Profile��5   	   $	��210��6   $	��559��6*   $��14(�66    � WindowStartupLocation$ CenterScreen=�6D    � 
ResizeMode$ NoResize=�5	   	   .��5
      � .+�5   
   �� $��Name:��6   $
��Left=�6    $��	10,10,0,0q�6;   $	��Top=�6N   �� 5   
   $��Associated Program:��6   $
��Left=�6.   $��	10,40,0,0q�6I   $	��Top=�5      �� 5   
   -   5         ��Name$ TBname��5      $
��Left=�6"   $��23��6=   $��
147,14,0,0q�6I   $	��Top=�5      $	��280��6*   �� 5   
   -   5        $ 	TBprogram��5      $
��Left=�6%   $��23��6@   $��
147,44,0,0q�6L   $	��Top=�5      $	��280��6*    �� 
IsReadOnly . 66   �� 5   
   -   5         ��Name$ Bbrowse��5   8   
���   �
���   �$��Browse��6I   $
��Left=�6Z   $��
433,44,0,0q�6u   $	��Top=�5      $	��100��6)   $��23��65   ��5      �� 5       �� EndPoint$ 0.5,1&�6&    � MappingMode$ RelativeToBoundingBox=�67    �� 
StartPoint$ 0.5,0&�6[   .��5      � 	 � Color$	 White��6#   � 5      $	 	#FFCDCDCD��6#   
 � Offset$
 1Q�65   5      5      �� 5   
   -   5         ��Name$ 
CHBenabled��5      $��Enabled��6'   $
��Left=�69   $��	15,78,0,0q�6T   $	��Top=�5      {�. 6+   �� 5   
   -   5        $	 Bok��5   8   
���   �
���   �$��OK��6E   $
��Left=�6R   $��
10,125,0,0q�6m   $	��Top=�6�   $	��100��5      $��30��6   ��5      �� 5       $ 0.5,1&�6&   $ RelativeToBoundingBox=�67   $ 0.5,0&�6[   .��5!      � $	 White��6#   � 5"      $	 	#FFCDCDCD��6#   $
 1Q�65   5#      5%      �� 5&   
   -   5        $ Bcancel��5&   8   
���   �
���   �$��Cancel��6I   $
��Left=�6Z   $��433,125,0,0q�6u   $	��Top=�5'      $	��100��6)   $��30��65   ��5(      �� 5)      $ 0.5,1&�6&   $ RelativeToBoundingBox=�67   $ 0.5,0&�6[   .��5*      � $	 White��6#   � 5+      $	 	#FFCDCDCD��6#   $
 1Q�65   5,      5.      5/      50      