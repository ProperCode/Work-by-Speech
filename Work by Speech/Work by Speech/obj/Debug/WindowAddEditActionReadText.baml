   M S B A M L   `   `   `  ����I  EWork by Speech, Version=2.2.0.0, Culture=neutral, PublicKeyToken=nullclr-namespace:SpeechSpeech  5   	   (    "Speech.WindowAddEditActionReadText   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   &localclr-namespace:Speech     5   	   -   5          � Title$!  Add Action: Read Aloud Text��5   	   $	��680��6-   $	��735��6:   $��14(�6F    � WindowStartupLocation$ CenterScreen=�5	   	    � 
ResizeMode$ NoResize=�6.   .��5      � .+�5   
   �� $��Text:��6   $
��Left=�6    $��5,5,0,0q�6;   $	��Top=�6L   �� 5   
   -   5         ��Name$ TBtext��5      $
��Left=�6"   $	��492��6=   $��	10,35,0,0q�6J    �� TextWrapping$
 Wrap=�5      $	��Top=�6&   $	��700��6>    �� AcceptsReturn . " �� VerticalScrollBarVisibility$ Visible=�5   '   �� 5   
   -   5         ��Name$ Btest��5   8   
���   �
���   �$
��Test��6G   $
��Left=�6V   $��
10,536,0,0q�6q   $	��Top=�5      $	��100��6)   $��30��65   ��5      �� 5       �� EndPoint$ 0.5,1&�6&   	 � MappingMode$	 RelativeToBoundingBox=�67   
 �� 
StartPoint$
 0.5,0&�6[   .��5      �  � Color$ White��6#   � 5      $ 	#FFCDCDCD��6#    � Offset$ 1Q�65   5      5      �� 5   
   -   5        $	 Bok��5   8   
���   �
���   �$��OK��6E   $
��Left=�6R   $��
10,598,0,0q�6m   $	��Top=�6�   $	��100��5      $��30��6   ��5      �� 5      $ 0.5,1&�6&   $	 RelativeToBoundingBox=�67   $
 0.5,0&�6[   .��5      � $ White��6#   � 5      $ 	#FFCDCDCD��6#   $ 1Q�65   5      5!      �� 5"   
   -   5        $ Bcancel��5"   8   
���   �
���   �$��Cancel��6I   $
��Left=�6Z   $��610,598,0,0q�6u   $	��Top=�5#      $	��100��6)   $��30��65   ��5$      �� 5%      $ 0.5,1&�6&   $	 RelativeToBoundingBox=�67   $
 0.5,0&�6[   .��5&      � $ White��6#   � 5'      $ 	#FFCDCDCD��6#   $ 1Q�65   5(      5*      5+      5,      