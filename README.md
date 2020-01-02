Weaver.py
===
```
Author: Rahul D. GHOSAL
Date: 1 January 2020
```
## 1. Purpose
Automatically generates a portion of a PowerPoint file used as a Simulation Report.
Table and other slide data are extracted from a designated Confirmation Tools file.

## 2. Usage
### 1. Installation
```bash
# To install as an editable package
pip install -e weaver 
```

### 2. Environment Settings
To use Weaver, a .txt textfile that lists paths to template PowerPoint files is necessary (refer to `paths_to_templates.txt`).
Please set the path to this file as an environment variable `TEMP_PATH` prior to executing this program.

### 3. 実行方法
```bash
# Help menu
weaver -h

# Execution
# NOTE: Windows-styled paths must be surrounded by quotation marks
weaver <Confirmation Tools PATH>
```

## 3. TODO
1. Implementing an algorithm that can take an input Simulation folder path and extract information about the ibis and buffer model of transmission line drivers and receivers.
2. Developing an algorithm to insert images into the appropriate slide (by e.g. using the image filename) 
3. Defining a ThermalReport class