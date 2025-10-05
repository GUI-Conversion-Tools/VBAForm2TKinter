# VBAForm2Tkinter - Excel VBA UserForm to Tkinter Converter
:jp:[日本語の説明はこちら](https://github.com/GUI-Conversion-Tools/VBAForm2Tkinter/blob/main/README_ja.md)<br><br>
This program converts userforms created in Microsoft Excel VBA into Python Tkinter code.<br>

## Example
<img width="681" height="1275" alt="Image" src="https://github.com/user-attachments/assets/27d11a87-58e0-4f6c-92ca-d1bbf69a8eae" /><br>
<img width="704" height="695" alt="Image" src="https://github.com/user-attachments/assets/ca514378-3017-443f-a3d8-bbd1ed4ceeb6" /><br>

## System Requirements
- Supported OS: Windows
- Required Software: Microsoft Excel

## Verified Operating Environments
= Windows 10/11
- Excel 2010(32bit)
- Excel 2016(32bit)
- Excel 2019(64bit)

## Converted Elements
- Variable names (object names)
- Approximate layout and size of controls
- Control colors (foreground, background)
- Text display (Label, CommandButton, CheckBox, ToggleButton, OptionButton, MultiPage)
- Font (typeface, size, bold, italic)
- Borders (UserForm, Frame, TextBox, Label, ListBox, Image)
- Mouse cursor
- Text alignment: left, center, right (Label, TextBox [MultiLine=False], ComboBox, CheckBox, ToggleButton, OptionButton)
- Default values of TextBox, ComboBox
- Items set in ComboBox, ListBox
- Selection state of OptionButton, CheckBox and ToggleButton
- Transparent background setting specified in BackStyle

## Supported Controls
| VBA Form Class | Tkinter Class|
| ------ | ------ |
| Label | tk.Label |
| CommandButton | tk.Button |
| Frame (without Caption) | tk.Frame |
| Frame (with any Caption) | tk.LabelFrame |
| TextBox (MultiLine=False) | tk.Entry |
| TextBox (MultiLine=True) | tk.Text |
| SpinButton | tk.Spinbox |
| ListBox | tk.Listbox |
| CheckBox | tk.Checkbutton |
| ToggleButton | tk.Checkbutton(indicatoron=0) |
| OptionButton | tk.Radiobutton |
| Image | tk.Canvas |
| ScrollBar | ttk.Scale |
| ComboBox | ttk.Combobox |
| MultiPage | ttk.Notebook |


> Note:
SpinButton behaves differently in VBA and Tkinter, so appearance may vary depending on placement.<br>
ScrollBar in VBA has up/down adjustment buttons, but Tkinter’s Scale does not.<br>
If unsupported controls exist on the form, the conversion will fail. If that case, please remove those controls and run the conversion again.<br>



## Usage
Before using, prepare the Excel workbook containing the user form you want to convert.
Also, ensure that the Immediate Window is visible in the VBE (Visual Basic Editor).<br><br>
<img width="843" height="768" alt="Image" src="https://github.com/user-attachments/assets/676cd54c-d610-4c25-bd9a-9e064e38dc5e" /><br><br>
1. Download the latest file from [here](https://github.com/GUI-Conversion-Tools/VBAForm2Tkinter/releases) and extract it. Use the VBAForm2Tkinter.bas file inside.<br>
2. In Excel, go to Developer -> Visual Basic to open VBE.<br>
3. Right-click your project and import the provided .bas file using Import File.<br>
4. In the Immediate Window, enter: Call ConvertForm2Tkinter(UserForm1)<br>
```vb
Call ConvertForm2Tkinter(UserForm1)
```
   > Note: Replace UserForm1 with the object name of the form you want to convert.

5.  If conversion succeeds, a message will appear, and an output.py file will be created in the same directory as your Excel workbook.<br>
6.  After checking the GUI appearance, edit the .py file and, above .mainloop(), configure event handlers for controls (e.g., button.configure(command=...)).<br>


## Control Order (for Controls Without Child Elements)
In Tkinter, if you place one Label on top of another, the later widget appears in front.<br>
However, in VBA, you can change front/back order, so the behavior differs.<br>
The program first sorts controls by hierarchy level; however, it preserves the original creation order within the same hierarchy.<br>
Since VBA’s z-order (front/back) cannot currently be retrieved, some displays may not match VBA.<br>

To adjust:<br>
&nbsp;&nbsp;&nbsp;&nbsp;Edit the Python code so the widget you want in front is placed later, or Reorder the controls in VBA before conversion.<br>
&nbsp;&nbsp;&nbsp;&nbsp;For new GUIs, instead of overlapping controls, it is recommended to use containers like Frame, which allow clear parent-child relationships.

## Notes on Usage
When using this program in a multi-monitor environment, please temporarily switch to a single monitor or ensure that all monitors have the same scaling percentage.
If monitors with different scaling percentages are mixed, the window size may not be calculated correctly.
