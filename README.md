# Delta2

Tool for finding and highlighting differences between excel files 

## Running as python file

1. Install requirements

```
pip install -r requirements.txt
```

2. Run the script

```
python run_delta2.py
```

## Assembling as a python bundle

1. Pack to run_delta2.exe

```
pyinstaller run_delta2.py --distpath InnoSetup/dist --noconsole --add-data="images;images"
```

2. Run run_delta2.exe from `InnoSetup\dist\run_delta2\run_delta2.exe`