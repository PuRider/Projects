   M S B A M L   `   `   `  ����T  PПрактика ТРПО, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullE'clr-namespace:Практика_ТРПОПрактика_ТРПО  5   	   -    'Практика_ТРПО.Authorization   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   9local'clr-namespace:Практика_ТРПО     5   	   b ^MaterialDesignThemes.Wpf, Version=4.0.0.2422, Culture=neutral, PublicKeyToken=df2a72020bd7962aFmaterialDesign1http://materialdesigninxaml.net/winfx/xaml/themes  5   	     � Title$  Аторизация��5	   	   $	��550��6   $	��400��6)    � 
ResizeMode$ NoResize=�65    � WindowStartupLocation$ CenterScreen=�6K   .��5
      � ��5   
   �� 5       �� EndPoint$ 0.5,1&�6"    �� 
StartPoint$ 0.5,0&�63   .��5      �  � Color$ 	#FF3F8797��6    � Offset$ 1Q�61   � 5      $ 	#FF3E3E3E��6   5      .+�5   
   �� $��Авторизация��6   $��40(�6&   $��Center=�64   $
��Left=�6X   $��	0,112,0,0q�6s   $	��Top=�6�   $��69��6�   $	��382��6�   ���	#DDFFFFFF$��Rockwell Extra Bold)�6�   �� 5   
   $��
Логин��6   $��24(�6    $
��Left=�6/   $��	0,273,0,0q�6J   $	��Top=�6]   $��35��6u   $	��125��6�   ���	#DDFFFFFF$��Rockwell Extra Bold)�6�   �� 5   
   $��Пароль��6   $��24(�6!   $
��Left=�60   $��	0,332,0,0q�6K   $	��Top=�6^   $��35��6v   $	��145��6�   ���	#DDFFFFFF$��Rockwell Extra Bold)�6�   �� 5   
   -   5         ��Name$ Login��5      )  #MaterialDesignThemes.Wpf.HintAssist   Hint$ admin��6!   .  (MaterialDesignThemes.Wpf.TextFieldAssist	   UnderlineBrush
	 �   �$
��Left=�6~   $��31��6�   
 �� TextWrapping$

 Wrap=�6�   $��161,273,0,0q�6�   $�� ��6�   $	��Top=�6�   $	��203��6�   $��16(�6�   
������� �� SelectionBrush
 �����$��0,0,0,2q�6H  
������� �� 
CaretBrush
 �������62  Z� 6v  �� 5   
   -   5        $ Password��5      $ admin��6$   
	 �   �$
��Left=�6�   $��31��6�   $

 Wrap=�6�   $��161,332,0,0q�6�   $�� ��6�   $	��Top=�6�   $	��203��6�   $��16(�6�   
�������
 �����$��0,0,0,2q�6K  
�������
 �������65  Z� 6y  �� 5   
   -   5         ��Name$ B1��5      $��
Войти��6   $
��Left=�6.   $��137.037,406.741,0,0q�6I   $	��Top=�6f   $��119.926��6~   $��42��6�   
�������$��20(�6�   
�������$��0q�6�    � UseLayoutRounding .  $��4q�6  $
��Bold"�60  $��Italic$�6B  ��6�   Z� ��5      �� 5       �� 
TargetType" ��6"   .R�5      ��  ��Name$ border��6   #��z��6-   #��z��6Y   #��z��6�    J� SnapsToDevicePixels .  �� CornerRadius$ 20��6�   ._�5      ��  ��Name$ contentPresenter��6+   #��z��6E   #��z��6y    �� ContentStringFormat �� ContentStringFormat# z 6�   |�.  #��z��6�   #��z��6.  ��. # z 6l  #��z��6�  5      5      5      5      5       