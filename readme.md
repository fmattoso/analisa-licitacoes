### Para Compilar
- No Windows, precisa ter instalado:
-- MSVC v143 - VS 2022 C++ x64/x86 build tools;
-- Windows 10/11 SDK;
-- C++ CMake tools for Windows;
-- Ou, se você tiver o Visual Studio instalado, instale o módulo "Desenvolvimento para Desktop com C++";\

Com o ambiente ativo (venv), insaler o Nuitka:\
> pip install nuitka\

E então execute a compilação:\
> python -m nuitka --standalone --onefile --windows-console-mode=disable --enable-plugin=pyside6 --include-data-file=produtos.db=produtos.db --output-filename=AnalisadorLicitacoes.exe main.py
