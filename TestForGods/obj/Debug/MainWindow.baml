   M S B A M L   `   `   `  ����F  BTestForGods, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null)clr-namespace:TestForGodsTestForGods  5   	       TestForGods.MainWindow   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   +localclr-namespace:TestForGods     5   	     � Title$  
MainWindow��5   	   $
��1366��6   $	��760��6)    � DataContext 5	         TestForGods.TestViewModel  5
   
   6     � 	Resources 5       �� MergedDictionaries �� Source�� 5   
    5      �� 5      $" Resources/MainResources.xaml?�6%   6M   %    5      .��5      �� $��Hidden=�6   $��Hidden=�6:   .��5   
   � $
��1366��6   $	��768��6   .+�5      � -   5         �Name$ TestPage��5      $y�Visible=�6%   .+�5      �� $��	0,20,0,40q�6   ���	#00000000._�5      � $��0,25,0,0q�6   $��Stretch=�6-    J� RenderTransformOrigin$ 1.455,0.529&�6I    � ColumnDefinitions 5      �� 5      �� 5      $
��1200 �6/   �� 5      .+�5      � $	��400��6   $	��Top=�6,   $��2��6D   $��
0,0,1000,0q�6X    � RowDefinitions 5      �� 5   "   �� 5    "   .+�5"      �� $��Images/logo.jpg��6$   $	��200��6=   $��0��6I   $	��200��6V   $��0,0,0,0q�6c   	 K� ZIndex$	 3��6t   $	��Top=�6�   6�   �� 5#      $��1��6)   
 �� Orientation$
 
Horizontal=�66   $��Bottom=�6O   $��Center=�6j   z�5$   "   �� 5%   &    �� ScaleX$
 0.65Q�65   6E   .+�5'   "   �� $,��&Fonts/Source Sans Pro/#Source Sans Pro)�6C      
CountStyle#��[  6w   $��50��6�   ��6,   �� *Index+6�   �� 5(   "   $��20,0,0,0q�6,   $�� /��6>   $,��&Fonts/Source Sans Pro/#Source Sans Pro)�6H   #��[  6|   $��60��6�   6�   �� 5)   "   $��20,0,0,0q�6,   $��10��6>   $,��&Fonts/Source Sans Pro/#Source Sans Pro)�6H   #��[  6|   $��80��6�   6�   5*      5+      � 5,      $��1��6   $��	100,0,0,0q�6/    5-      �� 5.   "   $	��300 �60   �� 5/   "   $
��Auto �60   �� 50   "   .+�52      � $	��300��6#   $	��Top=�60   $��	0,-30,0,0q�6H   .+�53   "   �� -   5         ��Name$ TestPicture��53   (   $	��300��6=   $��0,0,0,0q�6J   ��6[   �� *CurrentQuestion.ImageSource+6�   �� 54   "   $
��Bold"�6,   $��40(�6>   ��6L   �� *ViewWeather+6k   55      � 56      $	��190��6#   $��1��60   $��	0,-20,0,0q�6=   .+�57   "   �� $��Fonts/Scada/#Scada)�6,   $��Center=�58   ,   ���#000000$��Light"�5;   ,   $
��Wrap=�5<   ,   $
��Left=�5=   Z   $	��0,0q�6u   $
��1000��6�   ��5:   ,   �� *CurrentQuestion.Size+��5=   ,   �� *CurrentQuestion.QuestionText+z�5>   &   �� 5?   *    �� ScaleY$	 0.9Q�69   6H   5A   #   5B      � 5C      $��2��6#   $��0,0,0,30q�60    5D   "   �� 5E   &   �� 5F   &    5H   "   �� 5I   &   $��1* �67   �� 5J   &   $��1* �67   .+�5L   "   � $��0��6'   $��0��67   .+�5M   &   �� $
��Auto��6-   $��0��6U   ��6;   �� *Click+.��5N   *   � $
��Auto��6/    5O   .   �� 5P   2   $	��120 �6C   �� 5Q   2   .+�5S   .   A� $��75��66   $��75��6A   ���	#FFE84B43��6`   �� *	Thickness+��6�   �� *" CurrentQuestion.options[0].Color+6�   �� 5T   .   $��1��68   $
��Wrap=�6H   $
��Left=�6\   $��Fonts/Scada/#Scada)�6�   $��Center=�6�   $��40(�6�   ��6w   �� *" CurrentQuestion.options[0].Value+z�5U   2   �� 5V   6   $
 0.95Q�6E   6U   5X   /   5Y   +   5Z   '   5[   #   � 5\   "   $��1��6'   $��0��67   .+�5]   &   �� $
��Auto��6-   $��1��6U   ��6;   �� *Click+.��5^   *   � $
��Auto��6/    5_   .   �� 5`   2   $	��120 �6C   �� 5a   2   .+�5c   .   A� $��75��66   $��75��6A   ���	#FFE84B43��6`   �� *	Thickness+��6�   �� *" CurrentQuestion.options[1].Color+6�   �� 5d   .   $��1��68   $
��Wrap=�6H   $
��Left=�6\   $��Fonts/Scada/#Scada)�6�   $��Center=�6�   $��40(�6�   ��6w   �� *" CurrentQuestion.options[1].Value+z�5e   2   �� 5f   6   $
 0.95Q�6E   6U   5h   /   5i   +   5j   '   5k   #   � 5l   "   $��0��6'   $��1��67   .+�5m   &   �� $
��Auto��6-   $��2��6U   ��6;   �� *Click+.��5n   *   � $
��Auto��6/    5o   .   �� 5p   2   $	��120 �6C   �� 5q   2   .+�5s   .   A� $��75��66   $��75��6A   ���	#FFE84B43��6`   �� *	Thickness+��6�   �� *" CurrentQuestion.options[2].Color+6�   �� 5t   .   $��1��68   $
��Wrap=�6H   $
��Left=�6\   $��Fonts/Scada/#Scada)�6�   $��Center=�6�   $��40(�6�   ��6w   �� *" CurrentQuestion.options[2].Value+z�5u   2   �� 5v   6   $
 0.95Q�6E   6U   5x   /   5y   +   5z   '   5{   #   � 5|   "   $��1��6'   $��1��67   .+�5}   &   �� $
��Auto��6-   $��3��6U   ��6;   �� *Click+.��5~   *   � $
��Auto��6/    5   .   �� 5�   2   $	��120 �6C   �� 5�   2   .+�5�   .   A� $��75��66   $��75��6A   ���	#FFE84B43��6`   �� *	Thickness+��6�   �� *" CurrentQuestion.options[3].Color+6�   �� 5�   .   $��1��68   $
��Wrap=�6H   $
��Left=�6\   $��Fonts/Scada/#Scada)�6�   $��Center=�6�   $��40(�6�   ��6w   �� *" CurrentQuestion.options[3].Value+z�5�   2   �� 5�   6   $
 0.95Q�6E   6U   5�   /   5�   +   5�   '   5�   #   5�      5�      �� 5�      $��1��6!   $��2��6J   $��1100,0,30,0q�6^   ��61   �� *Next+.��5�      �� $��Images/button.jpg��6$   $	��120��6?   6M   5�      � 5�      $��3��6   $	��380��63   $
��Auto��6?   $��Bottom=�6M   $
��Left=�6h   $��	0,0,20,60q�6�   y�6�   �� *$"CurrentQuestion.options[0].Visible+.+�5�      �� $��Right=�6%   ���	#FF2E6C47$
��Auto��6X    �� CornerRadius$ 20��6e    �� Padding$ 20q�6w   $��0,0,40,0q�6�   $	��360��6�   ._�5�   "   �� $
��Wrap=�6,   $��22(�5�   ,   $��Fonts/Scada/#Scada)�5�   ,   ���#ffffff��5�   ,   ��  �� Path$, &CurrentQuestion.options[0].Explanation� �� Mode$ TwoWay=� �� UpdateSourceTrigger$ PropertyChanged=�5�   #   5�      � 5�       � Points �   @�� -1 Zb���	#FF2E6C47$��Right=�6S   6q   5�      � 5�      $��3��6   $	��400��63   $
��Auto��6?   $��Bottom=�6M   $
��Left=�6h   $��	20,0,0,60q�6�   y�6�   �� *$"CurrentQuestion.options[2].Visible+.+�5�      �� $��Right=�6%   ���	#FF2E6C47$
��Auto��6X   $ 20��6e   $ 20q�6w   $��0,0,40,0q�6�   $	��360��6�   ._�5�   "   �� $
��Wrap=�6,   $��22(�5�   ,   $��Fonts/Scada/#Scada)�5�   ,   ���#ffffff��5�   ,   �� $, &CurrentQuestion.options[2].Explanation�$ TwoWay=�$ PropertyChanged=�5�   #   5�      � 5�      % �    Zb@�� Zb -1 Zb���	#FF2E6C47$��Bottom=�6U   $��Right=�6p   6�   5�      � 5�      $��3��6   $	��380��63   $
��Auto��6?   $��Bottom=�6M   $��Right=�6h   $��	0,0,20,60q�6�   y�6�   �� *$"CurrentQuestion.options[1].Visible+.+�5�      �� $
��Left=�6%   ���	#FF2E6C47$
��Auto��6W   $ 20��6d   $ 20q�6v   $��40,0,0,0q�6�   $	��340��6�   ._�5�   "   �� $
��Wrap=�6,   $��22(�5�   ,   $��Fonts/Scada/#Scada)�5�   ,   ���#ffffff��5�   ,   �� $, &CurrentQuestion.options[1].Explanation�$ TwoWay=�$ PropertyChanged=�5�   #   5�      � 5�       �   @��@�� Zb���	#FF2E6C47$	��Top=�6S   $
��Left=�6k   6�   5�      � 5�      $��3��6   $	��380��63   $
��Auto��6?   $��Bottom=�6M   $��Right=�6h   $��	0,0,20,60q�6�   y�6�   �� *$"CurrentQuestion.options[3].Visible+.+�5�      �� $
��Left=�6%   ���	#FF2E6C47$
��Auto��6W   $ 20��6d   $ 20q�6v   $��40,0,0,0q�6�   $	��340��6�   ._�5�   "   �� $
��Wrap=�6,   $��22(�5�   ,   $��Fonts/Scada/#Scada)�5�   ,   ���#ffffff��5�   ,   �� $, &CurrentQuestion.options[3].Explanation�$ TwoWay=�$ PropertyChanged=�5�   #   5�      � 5�      % �    Zb@�� Zb@�� Zb���	#FF2E6C47$��Bottom=�6U   $
��Left=�6p   6�   5�      5�      5�      5�      � 5�      -   5        $ 	FinalPage��5�      ���#ffffff$��0,0,0,0q�6f   y�6&   �� *FinalPageVisibility+.+�5�      �� $��	0,20,0,40q�6   
���  ��$��0,0,0,0q�6G   ._�5�      � $��	0,12,0,88q�6   $��Stretch=�6.    5�      �� 5�      �� 5�      $
��1250 �6/   �� 5�      .+�5�      � $��1��6   .+�5�      �� $	��700��6)   $
 Vertical=�66   $��0,0,0,0q�6M   .+�5�   "   �� $��Center=�6-   $��Center=�6H   $
 
Horizontal=�6e   .+�5�   &   �� $
��Bold"�60   $��Fonts/Scada/#Scada)�6B   
���   �$��75(�6u   $��	Тест ��6�   6�   �� 5�   &   $
��Bold"�60   $��Fonts/Scada/#Scada)�6B   
���   �$�� завершён!��6u   $��75(�6�   6�   5�   #   �� 5�   "   $
 
Horizontal=�6-   $��Center=�6F   .+�5�   &   �� $��Bottom=�60   $��0,0,15,0q�6K   ���	#FFE84B43$,��&Fonts/Source Sans Pro/#Source Sans Pro)�6t   $��Regular"�6�   $	��140(�6�   ��6�   �� *CountTrueAnser+z�5�   *   �� 5�   .    �� CenterX$ 40Q�6=   6L   5�   '   �� 5�   &   ���	#FFE84B43$,��&Fonts/Source Sans Pro/#Source Sans Pro)�6H   $��Regular"�6|   $	��125(�6�   $��/��6�   6�   �� 5�   &   $��Bottom=�60   $��15,0,0,0q�6K   $,��&Fonts/Source Sans Pro/#Source Sans Pro)�6]   ���	#FFE84B43$��Regular"�6�   $	��125(�6�   $��10��6�   z�5�   *   �� 5�   .    �� Angle$ 1Q�6>    �� CenterX$ 0Q�6H   5�   '   5�   #   �� 5�   "   $
 
Horizontal=�6-   $
��1000��6F   .+�5�   &   �� $
��1000��60   $	��230��6=   $
��Wrap=�6J   $��37(�6^   ��6l   �� *ResultMessage+$ PropertyChanged=�$ TwoWay=�6�   5�   #   �� 5�   "   $��20��6)   67   �� 5�   "     BeginButton#��[ 6)   $	��400��6N   $	��100��6[   ��6h   �� *	NewGame+.��5�   &   �� $��Завершить��6,   $��55(�6@   $��Center=�6N   $��Medium"�6k   $��Center=�6   
�������6�   5�   #   5�      5�      5�      5�      5�      � 5�      -   5        $ 	FirstPage��5�      $��0,0,0,0q�6M   y�6&   �� *BeginVisibility+��5�      �� 5�       �� ImageSource$ Images/first.jpg��6!   6B   .+�5�      �� -   5        $	��448��5�      $	��105��6%   #��[ 6L   $��	0,400,0,0q�6q   .��5�      �� $��0,0,0,10q�6    $��Начать тест��62   ���#ffffff$��Center=�6Z   $��Center=�6w   $��Medium"�6�   $��	Open Sans)�6�   z�5�      �� 5�      $
 1.12Q�6-   6=   5�      5�      5�      5�      5�      5�      