pyinstaller --noconfirm --clean --onefile --windowed --name "DataTool" ^
    --exclude-module matplotlib ^
    --exclude-module scipy ^
    --exclude-module sklearn ^
    --exclude-module PIL ^
    --exclude-module tkinter.test ^
    --exclude-module unittest ^
    --exclude-module pydoc ^
    --exclude-module doctest ^
    --noupx ^
    data_tool_gui.py
