   M S B A M L   `   `   `  ����   TranslationAssistant%    TranslationAssistant.frmProfile   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml       � Title$  Profile Setting��$��501.253��$��345.432�� � WindowStyle$ 
ToolWindow=�$��Microsoft YaHei)�$
��11.5(� � 
ResizeMode$ NoResize=� � 	Resources    CenterGridHeaderStyle �� 
TargetType=  7System.Windows.Controls.Primitives.DataGridColumnHeader �� Property �� Value��%1   &	        �� "  .���  � ��$ Center=�.��� .+��� $��Parse Characters��$
��Left=�$��29��$��10,8,0,0q�$	��Top=�$	��111���� -    ��Name$ parseCharactersInput��$
��Left=�$��25��$��
126,12,0,0q� �� TextWrapping$
 Wrap=�$	��Top=�$	��166��$��、 。 ？ ！ 「 」��&   System.Windows.Media.TextOptions	   TextFormattingMode$	 Display=��� $��Profile��$
��Left=�$��23��$��	10,44,0,0q�$	��Top=�$	��101���� -   
 ��Name$
 	comboList��$
��Left=�$��25��$��
126,44,0,0q�$	��Top=�$	��166���� -    ��Name$ 	btnCreate��$��
Create New��$
��Left=�$��24��$��	69,88,0,0q�$	��Top=�$��80���� -   $ 	btnDelete��$��Delete Profile��$
��Left=�$��24��$��
179,88,0,0q�$	��Top=�$��80���� -   $ btnSave��$
��Save��$
��Left=�$��126,430,0,0q�$��80�� �� 	IsDefault . $��19��$	��Top=�&   System.Windows.Controls.DataGrid  -     Name$ 
dgvPhrases��$
��Left=�$	��298��$��
10,127,0,0q�$	��Top=�$	��304�� J� IsManipulationEnabled .    AutoGenerateColumns .     Columns 0  *System.Windows.Controls.DataGridTextColumn  ,  &System.Windows.Controls.DataGridColumn   Header$ Phrases��   Width5  /System.Windows.Controls.DataGridLengthConverter$
 0.5* 1  +System.Windows.Controls.DataGridBoundColumn   Binding �� *Phrase+  $ Translation��$
 0.5*  �� *Translation+