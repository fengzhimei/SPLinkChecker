   M S B A M L   `   `   `  ����   LinkChecker    LinkChecker.MainWindow   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml       � Title$  Broken Link Checker��$	��500��$	��705�� � 
ResizeMode$ CanMinimize=�.���  � RowDefinitions �� $	��120 ��� �� $��30 �.+�� $��0�� �� $��50 ��� $��40 ��� $��30 � � ColumnDefinitions �� $	��160 ��� .+��� $
��Left=�$��0��$��0��$��15(�$��	22,15,0,0q�$
��Wrap=�$��	Site URL:��$	��Top=��� $��1��$��0��$	��Top=�$
��Left=�$��0,10,0,0q�.+��� -   $��
txtSiteURL��$3��-http://sharepoint/sites/uknowlivedev/docstore��$	��Top=�$
��Left=�$��23�� �� TextWrapping$
 Wrap=�$	��510���� $
��Left=�$��0��$��1��$��15(�$��	22,13,0,8q�$
��Wrap=�$��Library Name:��$��Center=��� -   $��txtLibraryName��$
��Left=�$��
repository��$��Center=�$��24��$��1��$��1��$
 Wrap=�$	��510��$	��0,8q��� -   $��btnCheck��$��   Begin Scanning   ��$
��Left=�$��1��$��2��$��Center=�$��408,10,0,-11q�$��31��� $��1�� �� $��30 ��� $��35 ��� �� $��30 �.+��� -   $��spTimer��$y�Visible=�$��0��$��0�� �� Orientation$ 
Horizontal=�$��20,0,0,0q�$
��Left=�$��Center=�.+��� $��	Elapsed: ��
�������$��15(��� -    ��Name$ txtDuration��$��00:00��$��15(�$
��Bold"�
��������� -   $ 	txtStatus��$��10,0,0,0q�$��Complete��$��15(�$
��Bold"�
��� � �$y�	Collapsed=��� -   $��
svProgress��$��1��$��
20,10,20,0q�$��2��.���� 
�������$��1q�$��0��._��� -   $��txtProgress�� �� Padding$ 5q��� -	   $��sp��$��3��$ 
Horizontal=�.+��� $��Note: ��$
��Bold"�
���  ��$��20,0,0,0q�$��Center=�$��Center=��� -
   $��txtNote��$4��.You must have permission to access the library��$
��Bold"�$��Center=�$��Center=��� -    ��Name$ logSP��$��2��$��0��$ 
Horizontal=�$��20,0,0,0q�$y�	Collapsed=�.+��� $��Log file:  ��$
��Bold"�$��Center=�$��Center=��� $��Images/icxlsx.png��$��14��$��14��$��Center=�$��Center=��� -   $��LogLink��$��10,0,0,0q�$��Center=�$��Center=�