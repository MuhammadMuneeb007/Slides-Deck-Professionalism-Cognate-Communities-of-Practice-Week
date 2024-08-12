 
# Setup Instructions

## 1. Download Anaconda

Download Anaconda from the following link:

[Download Anaconda](https://www.anaconda.com/download)

## 2. Create and Activate Conda Environment

Run the following commands in your terminal to create and activate a new conda environment:

```bash
conda create -n academictools
conda activate academictools
```

## 3. Install Required Packages

Install the necessary Python packages using pip:

```bash
pip install python-pptx
pip install spellchecker
pip install pyspellchecker
pip install textstat
pip install pdf2docx PyPDF2
pip install docx2pdf
pip install imgkit
pip install enchant
pip install pyenchant
pip install LanguageTool
pip install language-tool-python
pip install gingerit
pip install deepgrammar
pip install deepgram-sdk
pip install grammar-check
pip install textblob
```
 
 
## 4. Prepare Your Files

Create a folder named `Files` and place all your slides in it.

The directory structure should look like this:

```
C:.
│   Cognate Communities of Practice Week.pptx
│   Installation.txt
│
└───Demo1
    │   Approach1-UsingPython.py
    │   Approach2-UsingPythonAndWindowsAPI.py
    │
    └───Files
            TECH1100 Week 01 Workshop.pptx
            TECH1100 Week 02 Workshop.pptx
            TECH1100 Week 03 Workshop.pptx
            TECH1100 Week 04 Workshop.pptx
            TECH1100 Week 05 Workshop.pptx
            TECH1100 Week 06 Workshop.pptx
            TECH1100 Week 07 Workshop.pptx
            TECH1100 Week 08 Workshop.pptx
            TECH1100 Week 09 Workshop.pptx
            TECH1100 Week 10 Workshop.pptx
            TECH1100 Week 11 Workshop.pptx
            TECH1100 Week 12 Workshop.pptx
```

Ensure the files are placed correctly as shown in the structure above.
 
            
## 5. Execute Approaches

### Approach 1: Using Python

- **Speed:** Fast
- **Limitations:** Limited options (SmartArt not processed)

Run the following command:

```bash
python Approach1-UsingPython
```
 

```bash
(academictools) C:\Users\kl\Demo1>python Approach1-UsingPython.py
```

### Output

The number of `.pptx` files in the directory `Files`: 12

#### File Details

|   No. | File Name                      | Author         | Last Modified By   | Created Date         | Modified Date        |
|------:|:-------------------------------|:---------------|:-------------------|:---------------------|:---------------------|
|     1 | TECH1100 Week 01 Workshop.pptx | Kathryn Cleary | Muhammad Muneeb    | 2015-04-27T02:22:59Z | 2024-03-18T22:04:16Z |
|     2 | TECH1100 Week 02 Workshop.pptx | Kathryn Cleary | Muhammad Muneeb    | 2015-04-27T02:22:59Z | 2024-03-25T23:44:16Z |
|     3 | TECH1100 Week 03 Workshop.pptx | Kathryn Cleary | Muhammad           | 2015-04-27T02:22:59Z | 2024-04-01T23:41:00Z |
|     4 | TECH1100 Week 04 Workshop.pptx | Kathryn Cleary | Muhammad Muneeb    | 2015-04-27T02:22:59Z | 2024-04-09T00:51:54Z |
|     5 | TECH1100 Week 05 Workshop.pptx | Kathryn Cleary | Muhammad Muneeb    | 2015-04-27T02:22:59Z | 2024-04-16T04:18:54Z |
|     6 | TECH1100 Week 06 Workshop.pptx | Kathryn Cleary | Djuro Mirkovic     | 2015-04-27T02:22:59Z | 2023-10-04T22:59:28Z |
|     7 | TECH1100 Week 07 Workshop.pptx | Kathryn Cleary | Muhammad Muneeb    | 2015-04-27T02:22:59Z | 2024-04-30T04:13:14Z |
|     8 | TECH1100 Week 08 Workshop.pptx | Kathryn Cleary | Muhammad Muneeb    | 2015-04-27T02:22:59Z | 2024-05-08T11:02:12Z |
|     9 | TECH1100 Week 09 Workshop.pptx | Kathryn Cleary | Muhammad Muneeb    | 2015-04-27T02:22:59Z | 2024-05-13T20:53:59Z |
|    10 | TECH1100 Week 10 Workshop.pptx | Kathryn Cleary | Muhammad           | 2015-04-27T02:22:59Z | 2024-05-20T22:28:14Z |
|    11 | TECH1100 Week 11 Workshop.pptx | Kathryn Cleary | Muhammad           | 2015-04-27T02:22:59Z | 2024-05-28T01:33:53Z |
|    12 | TECH1100 Week 12 Workshop.pptx | Kathryn Cleary | Djuro Mirkovic     | 2015-04-27T02:22:59Z | 2023-02-20T01:11:51Z |

#### Page Rendering Details

- **Loading page (1/2)**
- **Rendering (2/2)**
- **Done**

DataFrame converted to image and saved as `Files\TECH1100 Week 01 Workshop.pptx.png`

#### Slide Details

|    |   Slide Number | Heading Font Name   |   Heading Font Size | Text Font Name   | Text Font Size           |   Word count |   Animation count |   Bullet count |   Images count |
|---:|---------------:|:--------------------|--------------------:|:-----------------|:-------------------------|-------------:|------------------:|---------------:|---------------:|
|  0 |              1 |                     |                 nan | {'Arial'}        | {40.0, 32.0, 26.0, None} |           19 |                 0 |              0 |              0 |
|  1 |              2 |                     |                 nan | {'Arial', '-'}   | {16.0, 37.0}             |           47 |                 0 |              0 |              3 |
|  2 |              3 |                     |                 nan | {'Arial', '-'}   | {'-', 37.0}              |           19 |                 0 |              0 |              0 |
|  3 |              4 |                     |                 nan | {'Arial'}        | {40.0, 32.0, 26.0, None} |           19 |                 0 |              0 |              0 |
|  4 |              5 |                     |                 nan | {'Arial'}        | {14.77}                  |           69 |                 0 |              0 |              0 |
|  5 |              6 |                     |                 nan | {'-'}            | {9.0}                    |            5 |                 0 |              0 |              1 |
|  6 |              7 | -                   |                 nan | {'-'}            | {None}                   |            6 |                 0 |              0 |              0 |
|  7 |              8 |                     |                 nan | {'Arial', '-'}   | {24.0, 9.0, 37.0}        |           23 |                 0 |              0 |              1 |
|  8 |              9 |                     |                 nan | {'Arial', '-'}   | {24.0, 9.0, 37.0}        |           27 |                 0 |              0 |              1 |
|  9 |             10 |                     |                 nan | {'+mj-lt'}       | {44.0}                   |            3 |                 0 |              0 |              1 |
| 10 |             11 |                     |                 nan | {'Arial', '-'}   | {24.0, 9.0, 37.0}        |           19 |                 0 |              0 |              1 |
| 11 |             12 |                     |                 nan | {'Arial', '-'}   | {9.0, 20.0, 37.0}        |           46 |                 0 |              0 |              1 |
| 12 |             13 |                     |                 nan | {'Arial', '-'}   | {9.0, 20.0, 37.0}        |           25 |                 0 |              0 |              1 |
| 13 |             14 | -                   |                 nan | {'-'}            | {24.0, None}             |           85 |                 0 |              0 |              0 |
| 14 |             15 | -                   |                 nan | {'-'}            | {24.0, None}             |           22 |                 0 |              0 |              0 |
| 15 |             16 | -                   |                 nan | {'-'}            | {'-', None}              |            5 |                 0 |              0 |              0 |
| 16 |             17 | -                   |                 nan | {'-'}            | {24.0, None}             |           72 |                 0 |              1 |              0 |
| 17 |             18 | -                   |                 nan | {'-'}            | {24.0, None}             |          127 |                 0 |              0 |              0 |
| 18 |             19 | -                   |                 nan | {'-'}            | {24.0, 16.0, None}       |           31 |                 0 |              0 |              0 |
| 19 |             20 | -                   |                 nan | {'-'}            | {24.0, None}             |            7 |                 0 |              0 |              1 |
| 20 |             21 | -                   |                 nan | {'-'}            | {24.0, None}             |           64 |                 0 |              0 |              0 |
| 21 |             22 | -                   |                 nan | {'-'}            | {24.0, None}             |           82 |                 0 |              0 |              0 |
| 22 |             23 | -                   |                 nan | {'-'}            | {24.0, None}             |           23 |                 0 |              0 |              0 |
| 23 |             24 | -                   |                 nan | {'-'}            | {31.0, None}             |

           25 |                 0 |              0 |              0 |
| 24 |             25 | -                   |                 nan | {'-'}            | {24.0, None}             |           35 |                 0 |              0 |              0 |
| 25 |             26 | -                   |                 nan | {'-'}            | {31.0, None}             |           16 |                 0 |              0 |              0 |
| 26 |             27 | -                   |                 nan | {'-'}            | {24.0, None}             |           38 |                 0 |              0 |              0 |
| 27 |             28 | -                   |                 nan | {'-'}            | {24.0, None}             |           25 |                 0 |              0 |              0 |
| 28 |             29 | -                   |                 nan | {'-'}            | {24.0, None}             |           36 |                 0 |              0 |              0 |
| 29 |             30 | -                   |                 nan | {'-'}            | {24.0, None}             |           23 |                 0 |              0 |              0 |
| 30 |             31 | -                   |                 nan | {'-'}            | {24.0, None}             |           29 |                 0 |              0 |              0 |
| 31 |             32 | -                   |                 nan | {'-'}            | {24.0, None}             |           41 |                 0 |              0 |              0 |
| 32 |             33 | -                   |                 nan | {'-'}            | {24.0, None}             |           27 |                 0 |              0 |              0 |
| 33 |             34 | -                   |                 nan | {'-'}            | {24.0, None}             |           12 |                 0 |              0 |              0 |
| 34 |             35 | -                   |                 nan | {'-'}            | {24.0, None}             |           26 |                 0 |              0 |              0 |
| 35 |             36 | -                   |                 nan | {'-'}            | {24.0, None}             |           21 |                 0 |              0 |              0 |
| 36 |             37 | -                   |                 nan | {'-'}            | {24.0, None}             |           51 |                 0 |              0 |              0 |
| 37 |             38 | -                   |                 nan | {'-'}            | {24.0, None}             |           58 |                 0 |              0 |              0 |
| 38 |             39 | -                   |                 nan | {'-'}            | {24.0, None}             |           16 |                 0 |              0 |              0 |
| 39 |             40 | -                   |                 nan | {'-'}            | {24.0, None}             |           45 |                 0 |              0 |              0 |
```
 


### Approach 2: Using Python and Windows API

- **Speed:** Slow
- **Features:** Advanced features. Only processes words on Windows and requires COM library.

Run the following command:

```bash
python Approach2-UsingPythonAndWindowsAPI
```
```

You can save this content in a `.md` file and upload it to GitHub or include it in your project documentation.
