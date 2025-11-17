@echo off
REM Word to PowerPoint Converter - Стартиране
REM Двоен клик на този файл за да стартираш програмата

echo ================================================
echo   Word to PowerPoint Converter
echo ================================================
echo.
echo Стартиране на програмата...
echo.

REM Проверка дали Python е инсталиран
python --version >nul 2>&1
if errorlevel 1 (
    echo ГРЕШКА: Python не е намерен!
    echo.
    echo Моля инсталирай Python от:
    echo https://www.python.org/downloads/
    echo.
    echo Не забравяй да избереш "Add Python to PATH" при инсталация!
    echo.
    pause
    exit /b 1
)

REM Проверка за необходимите библиотеки
echo Проверка на библиотеките...
python -c "import docx, pptx, dateutil" >nul 2>&1
if errorlevel 1 (
    echo.
    echo Инсталиране на необходими библиотеки...
    echo.
    pip install -r requirements.txt
    if errorlevel 1 (
        echo.
        echo ГРЕШКА: Не успях да инсталирам библиотеките!
        echo Моля опитай ръчно: pip install -r requirements.txt
        echo.
        pause
        exit /b 1
    )
    echo.
    echo Библиотеките са инсталирани успешно!
    echo.
)

REM Стартиране на GUI апликацията
echo.
echo Стартиране на графичния интерфейс...
echo.
python word_to_ppt_gui.py

if errorlevel 1 (
    echo.
    echo Грешка при стартиране!
    pause
)
