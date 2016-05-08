@echo off

cd ..
for %%s in (*) do (
    python ./export/main.py %cd%/%%s
)

echo "-- done --------------"
pause
