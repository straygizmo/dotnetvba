# Manual test checklist

These checks open windows or dialogs, so they are not run by the automated
runner. Run each line in the Immediate window (Ctrl+G) of the VBE, once per host.
Access is untested by the maintainers; testers, please report your results.

## Excel

- [ ] `Net.Process.Start "notepad.exe"` -> Notepad opens.
- [ ] `Net.Process.Start Net.Environment.GetFolderPath(SpecialFolder_Desktop)` -> Explorer opens the Desktop.
- [ ] `?Net.Dialog.OpenFile("Pick", "Text", "*.txt")` -> dialog with a Text filter; cancel returns `""`.
- [ ] `?UBound(Net.Dialog.OpenFiles("Pick many"))` -> multi-select; choosing two files prints `1`.
- [ ] `?Net.Dialog.SaveFile("Save")` -> a SaveAs dialog opens.
- [ ] `?Net.Dialog.SelectFolder("Folder")` -> folder picker opens.

## Word

- [ ] `Net.Process.Start "notepad.exe"` -> Notepad opens.
- [ ] `Net.Process.Start Net.Environment.GetFolderPath(SpecialFolder_Desktop)` -> Explorer opens the Desktop.
- [ ] `?Net.Dialog.OpenFile("Pick", "Text", "*.txt")` -> dialog with a Text filter; cancel returns `""`.
- [ ] `?UBound(Net.Dialog.OpenFiles("Pick many"))` -> multi-select; choosing two files prints `1`.
- [ ] `?Net.Dialog.SaveFile("Save")` -> a SaveAs dialog opens.
- [ ] `?Net.Dialog.SelectFolder("Folder")` -> folder picker opens.

## Access (if available)

Access is untested by the maintainers. Please report what you observe.

- [ ] `Net.Process.Start "notepad.exe"` -> Notepad opens.
- [ ] `Net.Process.Start Net.Environment.GetFolderPath(SpecialFolder_Desktop)` -> Explorer opens the Desktop.
- [ ] `?Net.Dialog.OpenFile("Pick", "Text", "*.txt")` -> dialog with a Text filter; cancel returns `""`.
- [ ] `?UBound(Net.Dialog.OpenFiles("Pick many"))` -> multi-select; choosing two files prints `1`.
- [ ] `?Net.Dialog.SaveFile("Save")` -> `NotSupportedException` (error 1007) in Access.
- [ ] `?Net.Dialog.SelectFolder("Folder")` -> folder picker opens.
