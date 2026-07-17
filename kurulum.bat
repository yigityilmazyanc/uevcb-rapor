@echo off
rem ================================================================
rem  UEVCB Raporu - TEK SEFERLIK KURULUM (Windows)
rem  Cift tikla: paketleri kurar, ayar dosyasini olusturup acar,
rem  masaustune "UEVCB Raporu" kisayolu koyar.
rem  On kosul: Python 3.11+ (python.org - kurarken "Add python.exe
rem  to PATH" isaretli olmali). Gerisini bu dosya yapar.
rem ================================================================
chcp 65001 >nul
cd /d %~dp0

set PY=python
where py >nul 2>nul
if not errorlevel 1 set PY=py

%PY% --version >nul 2>nul
if errorlevel 1 goto python_yok

echo === [1/3] Gerekli paketler kuruluyor (eptr2, pandas, openpyxl) ===
%PY% -m pip install --quiet eptr2 pandas openpyxl
if errorlevel 1 goto pip_hata

echo.
echo === [2/3] Masaustu kisayolu ===
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0kisayol.ps1"

echo.
echo === [3/3] Ayar dosyasi ===
if exist ayarlar.json goto ayar_var
copy ayarlar.ornek.json ayarlar.json >nul
echo ayarlar.json olusturuldu - simdi Not Defteri'nde acilacak.
echo Iki alani doldurun ve KAYDEDIN:
echo   api_kullanici / api_sifre : EPIAS Seffaflik hesabi (kayit.epias.com.tr)
start /wait notepad ayarlar.json
goto son

:ayar_var
echo ayarlar.json zaten var - dokunulmadi.
goto son

:python_yok
echo HATA: Python bulunamadi.
echo https://www.python.org/downloads/ adresinden kurun;
echo kurulumda "Add python.exe to PATH" kutusunu isaretleyin, sonra bu dosyayi yeniden calistirin.
pause
exit /b 1

:pip_hata
echo HATA: paket kurulumu basarisiz - internet baglantisini kontrol edin.
pause
exit /b 1

:son
echo.
echo KURULUM TAMAM.
echo   Baslatma: masaustundeki "UEVCB Raporu" kisayolu (ya da arayuz.bat)
echo   Excel dosyalari Downloads klasorune yazilir.
pause
