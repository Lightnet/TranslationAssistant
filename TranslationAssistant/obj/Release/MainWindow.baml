   M S B A M L   `   `   `  ���� O  KTranslationAssistant, Version=2.0.0.0, Culture=neutral, PublicKeyToken=null;"clr-namespace:TranslationAssistantTranslationAssistant  %    TranslationAssistant.MainWindow   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     4local"clr-namespace:TranslationAssistant       � Icon$-  'pack://application:,,,/Resources/TA.ico�� � Title$  Translation Assistant Tool��$	��686��$	��566��$��Microsoft YaHei)�$
��12.5(� � WindowStyle$ ThreeDBorderWindow=�$	��500��$	��564��.��]� ��.  .+�o� -    o�Name$
 menu��$��23�� o� 
IsMainMenu . $	��top=�.@�m� $
��File��.@�m� -    m�Name$ menuNew��$��New (CTRL+N)��m� -   $ menuOpen��$��Open (CTRL+O)��m� -   $ menuSave��$��Save (CTRL+S)��{�.  m� $��Settings��.@�m� -   $ menuProfile��$��Profile (CTRL+P)��m� -   $ 
menuPhrase��$��Phrase (CTRL+L)��m� $��Special Punctuations��.@�m� -   $ 
menuSingle��$$��Single Qoute : 「　」  (F1)��m� -   $ 
menuDouble��$$��Double Qoute : 『　』  (F2)��m� -	   $ menuLenticular��$"��Lenticular : 【　】  (F3)��m� -
   $ menuEllipsis��$��Ellipsis : …  (F4)��m� -   $ menuWave��$��Wave Dash : 〜  (F5)��m� -   $ 
menuSTitle��$*��$Single Title Bracket : 〈 〉  (F6)��m� -   $ 
menuDTitle��$)��#Double Title Bracket : 《 》 (F7)��m� -   $ menuDash��$��Long Dash : ー (F8)��m� -   $ menuClipboard��$��Clipboard (CTRL+I)��{�.  m� -   $ 	menuAbout��$��About���� -    ��Name$ 	statusBar��$��23��$��Bottom=�$��Bottom=�.@��� -    ��Name$ completionStatus��$
��Left=�$��0% Complete���� -   $ 
LineStatus��$
��Left=�$��Line: xxxx/xxxx���� -   $ wordCountStatus��$
��Left=�$��
xxxx Words���� -   $ 	filesaved��$��Right=�$��
File Saved���� -    ��Name$ 	reviewTop��	 �� TextWrapping$
	 Wrap=�$����HOW TO USE:
Creating New File for Translations:
1.)Click File->New
2.)Copy Raw from Source and paste it on the textbox
3.)Click Create
4.)Save File into the desired location
5.)Translate
6.)Once done translating click on Clipboard. Paste on your preferred editor or your blog for final edit and posting

Creating New Profile
1.)Click Settings->Profile
2.)Click New Profile
3.)Type profile name
4.)Click OK

Adding Phrases to Profile:
Option 1:
1.)Click Settings->Add Phrase
2.)Enter Raw on Phrase textbox
3.)Enter translation of phrase in Translation textbox
4.)Click Save
Option 2:
1.)Click Settings->Profile
2.)Add Raw text and Translated Test on their respective boxes at the end of table
3.)Click Save

Deleting Phrases
1.)Click Settings->Profile
2.)Double-Click phrase you want to delete
3.)Confirm Delete
4.)Click Save

Editing Phrases
1.)Click Settings->Profile
2.)Click on a phrase, and click once again (make sure to have a pause between clicks or else it will be considered a double click and will ask to delete an entry) 
3.)Modify entry
4.)Click Save

Use Spelling Suggestions
1.)Right-Click the red underlined word
2.)Choose the correct spelling you meant to type from the context menu

Adding to Custom Dictionary
1.)Highlight word to add to custom dictionary (Double click on the word)
2.)Press CTRL+J and confirm the action or Right-Click and Add To Dictionary

Special Punctuations
1.) You can click on the Special Punctuations menu to insert the desired punctation or press the corresponding Function key (F1 - F5)��$	��Top=�$��5q�$	��320��
 �� 
IsReadOnly
 . " �� VerticalScrollBarVisibility$ Visible=��� -   $ currentRawLine��$��52��$
	 Wrap=�$��5q�$	��Top=�
 . $ Visible=� �� IsUndoEnabled .  ]� $	��Top=� +� IsFocusScope . .+��� -   $ currentTranslatedLine��$��52��$
	 Wrap=�$��5q�$	��Top=� �� 	IsEnabled . $ Visible=�.�� � ContextMenu �� .@�m� .@�dummy.���� -   $ reviewBottom��$	��137��$
	 Wrap=�$��5q�$	��Top=�
 . $ Visible=�$����NAVIGATION CONTROLS:
ENTER Key or PgDn Key  = Move down next sentence.
PgUp Key = Move up to previous sentence
CTRL+Left Key = Highlight previous parsed phrase in a sentence, if this was the first parsed phrase it will revert the sentence to original raw if there was any glossary swapping done
CTRL+Right Key = Highlight next parsed phrase in a sentence, if this was the original raw that was reverted from a glossary swap it will swap any glossary terms
CTRL+Home Key = Go to the the first line
CTRL+End Key = Go to the most recent un-translated sentence��* �� #IsInactiveSelectionHighlightEnabled . 