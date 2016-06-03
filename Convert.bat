for /R %%I in (*.xlsx) do (
    xlsx2csv.exe -f %%I
)
pause
