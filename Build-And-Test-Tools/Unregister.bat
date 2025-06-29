:: Change directory up 1 then into the release directory
cd "..\target\x86_64-pc-windows-msvc\release"

:: UnRegister the DLL
regsvr32 /u win_svg_thumbs.dll
