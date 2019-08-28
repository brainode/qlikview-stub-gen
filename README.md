# qlikview-stub-gen
Stub generator for QlikView.

Generate "stubs"(Text objects) over Charts. At this early moment works **only**, when chart id starts from CH.

Example of usage:
```
StubGenerator.exe -p "C:\Qv\Test_Project-prj" -z 20 -t "No data" --bgcolor RGB(255,255,255)
```
Where:
* p - is path to prj folder with qvw xml files.
* z - is Z-level on sheet.
* t - is text on Stub
* bgcolor - is stub background color. Variable name can be placed instead RGB
